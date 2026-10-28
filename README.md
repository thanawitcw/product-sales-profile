# Product Sales Profile (R)

Builds a clean, promo-attributed daily sell-out dataset and a per-product sales
profile from raw sell-out exports. Every sales row is mapped to a sales type
(`Normal` / `Promotion`), a promotion theme, and an effective selling price
including VAT, then rolled up into average quantity and carton sales per day.

This is the consolidated version of two earlier R scripts (daily sell-out + promo
cleaning and product sales profiling). The simpler single-month script is
superseded by this one and has been removed to keep a single R project.

---

## What it does

1. Loads the latest N monthly Excel files from the sell-out folder. The month is
   parsed from the filename (`YYYY_M...`), so files do not need to arrive in order.
2. Loads the latest N monthly TD sell-out CSV files and appends them as a second
   sales channel.
3. Reads the product master from a Microsoft Access database to get:
   - `CJ_Item` values flagged `non_vat = TRUE`
   - `PC_Cartons` per item (discontinued items excluded)
4. Cleans the totals: negative or missing quantities/amounts become `0`, and rows
   are aggregated by `Date + CJ_Item`.
5. Adds **7% VAT** to amount fields, except for items flagged non-VAT.
6. Computes the average normal selling price:
   `avg_normal_price = round_excel(sum_normal_amt / sum_normal_qty)`.
7. Summarises the promotion detail (008) by
   `Date, CJ_Item, Promotion_Type, Promotion_Theme, PromoPrice In.Vat`.
   Themes containing `"แถม"` (free-gift / BOGO) derive their price from
   `promotion amount / promotion quantity`.
8. Allocates the daily totals across promotion rows proportionally, so the
   promotion split always adds back to the original total.
9. Marks `Main_Promotion = "Main"` when one promotion is more than 50% of the
   item's quantity for that day.
10. Builds the **Sales_Profile** sheet: average quantity per day and average
    cartons per day for each item × promotion type × theme × price.
11. Writes both the raw daily dataset and the profile to one Excel workbook.

---

## Data sources

Paths are configurable through environment variables; the defaults are relative
to this folder.

| Input | Env var | Default | Format |
|---|---|---|---|
| Daily sell-out totals (001) | `DAILY_SALES_PATH` | `data/raw/Daily_Sellout_001` | One `.xlsx` per month |
| Promotion detail (008) | `DAILY_PROMOTION_PATH` | `data/raw/Daily_Sellout_008` | One `.xlsx` per month |
| TD sell-out | `TD_SALES_PATH` | `data/raw/TD_SellOut` | `.csv` per month |
| Product master | `ACCESS_DB_PATH` | `data/raw/master.accdb` | Access table `tbl_master_003_product_list` |
| Output folder | `SALES_OUTPUT_DIR` / `SALES_OUTPUT_FILE` | `data/output/RawData_sales_profile.xlsx` | `.xlsx` |

Root override: `SALES_DATA_ROOT` (default `data/raw`).

Expected columns in the sell-out 001 file (by position 1, 5, 6, 14–17):
`Calendar Day` (parsed as `DD.MM.YYYY`), `Product` (mapped to `CJ_Item`),
the product name column, `Normal Sale/Base QTY (TY)`,
`Normal Sale/AMT (TY) Exc.VAT`, `Promotion Sale/Base QTY (TY)`,
`Promotion Sale/AMT (TY) Exc.VAT`.

Expected columns in the promotion 008 file: `Date`, `CJ_Item`,
`Promotion Base QTY (TY)`, `Promotion Amount TY`, `Promotion_Type`,
`Final_Pro_Theme`, `PromoPrice In.Vat`.

---

## Output

`RawData_sales_profile.xlsx`

| Sheet | Grain | Key columns |
|---|---|---|
| `Raw_Daily_SO` | date × item × sales type × promo | `Date`, `CJ_Item`, `Product_Name`, `Sales Channel`, `Promotion_Type`, `Promotion_Theme`, `ProPrice In.Vat`, `Total_SO_Qty`, `Total_SO_Amt (In.Vat)`, `Main_Promotion` |
| `Sales_Profile` | item × promo | `CJ_Item`, `Product_Name`, `Promotion_Type`, `Promotion_Theme`, `ProPrice In.Vat`, `avg_qty_per_day_from_CJ`, `avg_ctn_per_day_from_CJ`, `PC_Cartons` |

---

## Requirements

- R 4.x
- Packages: `readxl`, `dplyr`, `tidyr`, `readr`, `writexl`, `stringr`, `RODBC`
- Microsoft Access Database Engine (x64) if you read the Access product master

```r
install.packages(c("readxl", "dplyr", "tidyr", "readr", "writexl", "stringr", "RODBC"))
```

## Run

```r
source("Product_Sales_Profile.R")
```

Or from a terminal:

```bash
Rscript Product_Sales_Profile.R
```

Before running, point the env vars at your own folders (or place the files under
`data/raw/` using the table above). Set `num_files` on the two
`load_and_process_files()` calls if you want a different history window.

---

## Notes

- `round_excel()` reproduces Excel's half-up rounding instead of R's default
  banker's rounding, so totals match the source reports.
- The script prints a reminder that `num_files` in the TD loader controls how many
  months are appended.
- All company data, credentials, and internal paths have been removed from this
  repository. Use your own inputs; do not commit business data.
