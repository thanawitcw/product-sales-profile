library(readxl)
library(dplyr)
library(tidyr)
library(readr)
library(writexl)
library(stringr)
library(RODBC)

# Define file paths as constants.
# Override any of these with environment variables; defaults are relative to
# this project folder so the script runs on any machine.
data_root <- Sys.getenv("SALES_DATA_ROOT", "data/raw")

daily_sales_path <- Sys.getenv(
    "DAILY_SALES_PATH", file.path(data_root, "Daily_Sellout_001"))
daily_promotion_path <- Sys.getenv(
    "DAILY_PROMOTION_PATH", file.path(data_root, "Daily_Sellout_008"))
td_sales_path <- Sys.getenv(
    "TD_SALES_PATH", file.path(data_root, "TD_SellOut"))

access_path <- Sys.getenv(
    "ACCESS_DB_PATH", file.path(data_root, "master.accdb"))


# Function to load and process Excel files
load_and_process_files <- function(file_path, num_files = 5, file_type = "auto") {
    # Determine file extension based on file type
    if (file_type == "excel"){
        pattern <- "\\.xlsx$"
    }
    else if (file_type == "csv"){
        pattern <- "\\.csv$"
    }
    else {
        # check what file extension in the directory
        excel_files <- list.files(file_path, pattern = "\\.xlsx$", full.names = TRUE)
        csv_files <- list.files(file_path, pattern = "\\.csv$", full.names = TRUE)
        if (length(excel_files) >0) {
            pattern <- "\\.xlsx$"
            file_type <- "excel" 
        } else if (length(csv_files) >0) {
            pattern <- "\\.csv$"
            file_type <- "csv"
        } else {
            stop("No Excel or CSV file found in the directory")
        }
    }
    
    files <- list.files(file_path, pattern = pattern, full.names = TRUE)
    files <- files[!grepl("^~\\$", basename(files))]
    date_pattern <- "^([0-9]{4})_([0-9]{1,2})"
    
    dates <- sub(date_pattern, "\\1-\\2-01", basename(files))
    dates <- as.Date(dates, format = "%Y-%m-%d")
     
    sorted_files <- files[order(dates)]
    selected_files <- tail(sorted_files, num_files)
    
    data_list <- list()
    
    for (file in selected_files) {
        tryCatch({
            if (file_type == "excel"){
                df <- read_excel(file)
            } else {
                df <- read_csv(file, locale = locale(encoding = "UTF-8"), show_col_types = FALSE)
            }
            data_list[[file]] <- df
            print(paste("Loaded:", file))
        }, error = function(e) {
            print(paste("Error loading:", file, "-", e$message))
        })
    }
    
    return(data_list)
}


# Define a round formula for R (ignore Banker Method -> always round to the closet even number)
round_excel <- function(x) {
    sign(x) * floor(abs(x) + 0.5)
}


# Connect to Microsoft Access Database and fetch non-Vat items
conn <- odbcConnectAccess2007(access_path)
qry_non_vat_items <- "SELECT DISTINCT CJ_Item
                      FROM tbl_master_003_product_list
                      WHERE non_vat = True"

non_vat_items <- sqlQuery(conn, qry_non_vat_items)

qry_data_pc_cartons <- "SELECT CJ_Item,
                               PC_Cartons
                        FROM tbl_master_003_product_list
                        WHERE [Group] <>'Discontinuous'"

data_pc_cartons <- sqlQuery(conn, qry_data_pc_cartons)

odbcClose(conn)


# Create function to include VAT 7% and skip for the non_vat_items list
include_vat <- function(df, column_name, non_vat_items) {
    for (col in column_name) {
        df <- df %>% 
            mutate(!!sym(col) := if_else(!(CJ_Item %in% non_vat_items$CJ_Item),
                                         !!sym(col) * 1.07,
                                         !!sym(col)))
    }
    return(df)
}

# Load sales and promotion data
total_sales_list <- load_and_process_files(daily_sales_path, file_type = "excel")
promotion_sales_list <- load_and_process_files(daily_promotion_path, file_type = "excel")

# Combine the lists of data frames into single data frames
total_sales_df <- bind_rows(total_sales_list)

# Pivot to get 1 Product name per CJ_Item
product_name_reference <- total_sales_df %>% 
    select(5, 6) %>% 
    rename(
        CJ_Item = Product,
        Product_Name = names(.)[2]
    ) %>% 
    group_by(CJ_Item) %>% 
    slice_tail(n=1) %>% 
    ungroup()

# Construct Data Transform
total_sales_df <- total_sales_df %>%
    # Select necessary column
    select(1, 5, 14, 15, 16, 17) %>%
    # Rename Column for further step
    rename(
        Date = `Calendar Day`,
        CJ_Item = Product,
        Normal_Qty = `Normal Sale/Base QTY (TY)`,
        Normal_Amt = `Normal Sale/AMT (TY) Exc.VAT`,
        Promotion_Qty = `Promotion Sale/Base QTY (TY)`,
        Promotion_Amt = `Promotion Sale/AMT (TY) Exc.VAT`
    ) %>% 
    # Replace Null and <0 values
    mutate(
        across(c(Normal_Qty, Normal_Amt, Promotion_Qty, Promotion_Amt), ~pmax(replace_na(., 0),0))
    ) %>%
    # Include Vat to Amt column
    include_vat(c("Normal_Amt", "Promotion_Amt"), non_vat_items) %>%
    # Group by data
    group_by(Date, CJ_Item) %>% 
    summarise(
        across(c(Normal_Qty, Normal_Amt, Promotion_Qty, Promotion_Amt), sum), .groups = "drop"
    ) %>%
    # Unpivot column
    pivot_longer(
        cols = c(Normal_Qty, Normal_Amt, Promotion_Qty, Promotion_Amt),
        names_to = "Attribute",
        values_to = "Value"
    ) %>% 
    # split column name to pivot wider
    separate(Attribute, into = c("Sales_Type", "Sales_Metrics"), sep = "_", remove =TRUE) %>% 
    # Pivot wider
    pivot_wider(
        names_from = Sales_Metrics,
        values_from = Value
    ) %>% 
    # Change Data Types
    mutate(
        CJ_Item = as.character(CJ_Item),
        Sales_Type = as.character(Sales_Type),
        Date = as.Date(Date, format = "%d.%m.%Y")
    ) %>%
    # Arrange Date
    arrange(Date) %>%
    # merge DF to get product name
    left_join(product_name_reference, by = "CJ_Item")

## Transform data to get ASP for normal sales
normal_RSP <- total_sales_df %>%
    select(CJ_Item, Sales_Type, Qty, Amt) %>%
    filter(Sales_Type == "Normal") %>% 
    group_by(CJ_Item, Sales_Type) %>% 
    summarise(
        sum_normal_qty = sum(Qty, na.rm = TRUE),
        sum_normal_amt = sum(Amt, na.rm = TRUE),
        .groups = "drop") %>% 
    mutate(
        avg_normal_price = round_excel(sum_normal_amt / sum_normal_qty)
    ) %>%
    # Select only 2 columns
    select(CJ_Item, Sales_Type ,avg_normal_price)

# Merge with original dataframe to get AVG normal price
total_sales_df <- left_join(total_sales_df, normal_RSP, by = c("CJ_Item", "Sales_Type")) %>% 
    mutate(avg_normal_price = if_else(Qty == 0, 0, avg_normal_price))


# working with Promotion sales Report
promotion_sales_df <- bind_rows(promotion_sales_list)
# construct data pipeline
promotion_sales_df <- promotion_sales_df %>% 
    select(
        `Calendar Day`,
        Material,
        `Promotion Base QTY (TY)`,
        `Promotion Amount TY`,
        Promotion_Type,
        Final_Pro_Theme,
        `Final_Price/Unit (In.Vat)`
    ) %>% 
    rename(
        Date = `Calendar Day`,
        CJ_Item = Material,
        Promotion_Theme = Final_Pro_Theme,
        `PromoPrice In.Vat` = `Final_Price/Unit (In.Vat)`
    ) %>%
    include_vat("Promotion Amount TY", non_vat_items)


# Handle Promotion_Theme that contains "แถม" to get average RSP
bogo_ASP <- promotion_sales_df %>%
    # filter data to get average RSP
    filter(grepl("แถม", Promotion_Theme)) %>%
    group_by(CJ_Item, Promotion_Type, Promotion_Theme) %>%
    summarise(
        sum_promo_qty = sum(`Promotion Base QTY (TY)`, na.rm = TRUE),
        sum_promo_amt = sum(`Promotion Amount TY`, na.rm = TRUE),
        .groups = "drop"
    ) %>%
    mutate(
        ASP_from_bogo = round_excel(sum_promo_amt / sum_promo_qty)
    )


# Handle other Promotion_Theme
promotion_sales_df <- promotion_sales_df %>%
    group_by(Date, CJ_Item, Promotion_Type, Promotion_Theme, `PromoPrice In.Vat`) %>%
    summarise(
        Promotion008_Qty = sum(`Promotion Base QTY (TY)`),
        `Promotion008_Amt (In.Vat)` = sum(`Promotion Amount TY`),
        .groups = "drop"
    ) %>% 
    mutate(
        Date = as.Date(`Date`, format ="%d.%m.%Y"),
        Sales_Type = "Promotion"
    )

# merge to total sales report to prep promotion sales
total_sales_df <- total_sales_df %>%
    left_join(promotion_sales_df,
              by = c("Date", "CJ_Item", "Sales_Type")) %>%
    
    ## merge with bogo_promotion to get average RSP
    left_join(bogo_ASP,
              by = c("CJ_Item", "Promotion_Type", "Promotion_Theme")) %>%
    mutate(
        `ProPrice In.Vat` = case_when(
            Sales_Type == "Normal" ~avg_normal_price,
            str_detect(Promotion_Theme, "แถม") ~ ASP_from_bogo,
            TRUE ~ `PromoPrice In.Vat`
        )
    ) %>% 
    filter(!(Sales_Type == "Promotion" & is.na(Promotion008_Qty))) %>% 
    mutate(
        Promotion_Type = replace_na(Promotion_Type, "Normal"),
        Promotion_Theme = replace_na(Promotion_Theme, "Normal Price")
    )

# Prepare Data by promotion for convert
sum_sales <- total_sales_df %>% 
    group_by(Date, CJ_Item, Sales_Type) %>%
    summarise(
        sum_promotion008_qty = sum(Promotion008_Qty, na.rm = TRUE),
        sum_promotion008_amt = sum(`Promotion008_Amt (In.Vat)`),
        .groups = "drop"
    )

# Convert sales data to be actual
total_sales_df <- total_sales_df %>%
    left_join(sum_sales, by = c("Date", "CJ_Item", "Sales_Type")) %>%
    mutate(
        Total_SO_Qty = if_else(
            Sales_Type == "Normal",
            Qty,
            round_excel((Promotion008_Qty / sum_promotion008_qty) * Qty)
        ),
        `Total_SO_Amt (In.Vat)` = if_else(
            Sales_Type == "Normal",
            Amt,
            (`Promotion008_Amt (In.Vat)` / sum_promotion008_amt) * Amt
        )
    ) %>%
    # Replace Error value on Total_SO_Amt In.Vat
    mutate(
        `Total_SO_Amt (In.Vat)` = replace_na(`Total_SO_Amt (In.Vat)`, 0)
    ) %>% 
    # Drop unnecessary columns
    select(
        Date,
        CJ_Item,
        Product_Name,
        Promotion_Type,
        Promotion_Theme,
        `ProPrice In.Vat`,
        Total_SO_Qty,
        `Total_SO_Amt (In.Vat)`
    ) %>%
    # Summarize data for same promotion has multiple rows
    group_by(Date,
             CJ_Item,
             Product_Name,
             Promotion_Type,
             Promotion_Theme,
             `ProPrice In.Vat`) %>% 
    summarise(Total_SO_Qty = sum(Total_SO_Qty),
              `Total_SO_Amt (In.Vat)` = sum(`Total_SO_Amt (In.Vat)`),
              .groups = "drop") %>%
    # Group by data to summary total sales per day
    group_by(Date, CJ_Item) %>% 
    mutate(
        Main_Promotion = if_else(
            sum(Total_SO_Qty) > 0 & Total_SO_Qty / sum(Total_SO_Qty) > 0.5,
            "Main",
            "-"
        )
    ) %>% 
    ungroup()


# Transform the data to deal with multiple CJ_Item
pc_cartons <- data_pc_cartons %>%
    mutate(CJ_Item = as.character(CJ_Item)) %>% 
    group_by(CJ_Item) %>% 
    summarize(PC_Cartons = max(PC_Cartons))


## working with product sales profile
product_sales_profile <- total_sales_df %>% 
    filter(Main_Promotion == "Main",
           Total_SO_Qty > 0) %>%
    group_by(CJ_Item, 
             Product_Name, 
             Promotion_Type, 
             Promotion_Theme, 
             `ProPrice In.Vat`) %>%
    summarise(
        Total_SO_Qty = sum(Total_SO_Qty),
        Trans_sales = n(),
        .groups = "drop"
    ) %>% 
    mutate(
        avg_qty_per_day_from_CJ = Total_SO_Qty / Trans_sales
    ) %>%
    filter(!is.na(avg_qty_per_day_from_CJ)) %>%
    select(
        -Total_SO_Qty,
        -Trans_sales
    )

# Merge to product sales profile to convert avg qty to avg ctn
product_sales_profile <- product_sales_profile %>%
    left_join(pc_cartons, by = "CJ_Item") %>%
    mutate(avg_ctn_per_day_from_CJ = if_else(PC_Cartons > 0,
                                             avg_qty_per_day_from_CJ / PC_Cartons,
                                             NA_real_)) %>% 
    select(
        CJ_Item,
        Product_Name,
        Promotion_Type,
        Promotion_Theme,
        PC_Cartons,
        everything()
    )

## Load TD sales and append to total_sales_df
sales_td_lists <- load_and_process_files(td_sales_path, num_files = 3, file_type = "csv")
## Adding Alert message
message("Don't forget to change num_files parameter in sales_td_lists")
## Append all file
sales_td <- bind_rows(sales_td_lists)
## Construct Data pipeline
sales_td <- sales_td %>%
    mutate(CJ_Item = as.character(CJ_Item),
           Date = as.Date(Date, format = "%m/%d/%Y")) %>%
    filter(Main_Promotion == "Main") %>% 
    left_join(product_name_reference, by = "CJ_Item")


## Appended sales CJ + TD together
final_total_sales_df <- bind_rows(total_sales_df, sales_td) %>%
    filter(!(Total_SO_Qty == 0)) %>%
    arrange(Date) %>%
    mutate(`Sales Channel` = coalesce(`Sales Channel`, "CJ")) %>% 
    select(
        Date,
        CJ_Item,
        Product_Name,
        `Sales Channel`,
        everything()
    )


# Define save path (override with SALES_OUTPUT_FILE)
output_dir <- Sys.getenv("SALES_OUTPUT_DIR", "data/output")
dir.create(output_dir, recursive = TRUE, showWarnings = FALSE)
output_file_path <- Sys.getenv(
    "SALES_OUTPUT_FILE", file.path(output_dir, "RawData_sales_profile.xlsx"))
# list the data to save in excel file
excel_data <- list(
    "Raw_Daily_SO" = final_total_sales_df,
    "Sales_Profile" = product_sales_profile)

# save to excel file
write_xlsx(excel_data, output_file_path)

