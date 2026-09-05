# Data Cleaning & Reporting Automation

A reusable Python tool that cleans messy tabular data (CSV/Excel) and automatically generates a
data-quality report (HTML, with charts) — so cleaning a new export becomes a single command instead
of a manual spreadsheet slog.

## What it does

Given a raw CSV/Excel file, it:
1. **Profiles** the data (row count, missing values, duplicates, completeness %) — before and after
2. **Cleans** it:
   - Removes exact duplicate rows and duplicate keys (e.g. repeated Order IDs)
   - Trims whitespace and standardizes text casing (names, categories)
   - Standardizes inconsistent category spellings ("north"/"NORTH"/"N." → "North") — both via an
     explicit mapping you supply and automatic grouping of near-identical values
   - Parses mixed date formats into one consistent ISO format
   - Strips currency symbols/commas from price fields and converts them to numbers
   - Fixes sign errors (negative quantities) and statistical outliers (IQR method)
   - Fills missing values per-column using a configurable strategy (median/mean/mode/flag/drop)
   - Flags structurally invalid values (e.g. malformed emails) for manual review, rather than
     guessing a "fix" that could be wrong
3. **Generates an HTML report** — KPI summary, before/after charts, a full log of every action taken
   (with counts), and a preview of the cleaned data
4. **Saves the cleaned file** as CSV or Excel

## Files
- `data_cleaning_automation.py` — the reusable engine (also works as a CLI tool)
- `generate_messy_data.py` — creates the demo dataset (868 rows of realistic messy order data)
- `messy_orders.csv` — the demo input
- `cleaned_orders.csv` — the demo output
- `cleaning_report.html` — the demo generated report (open in any browser)

## Running it on the demo data
```bash
python data_cleaning_automation.py messy_orders.csv \
    --output-data cleaned_orders.csv \
    --output-report cleaning_report.html
```

## Running it on YOUR data
The engine is dataset-agnostic — only a small config tells it which columns need which treatment.
Write a JSON config (see `DEMO_CONFIG` in the script for the exact shape) describing:
- `duplicate_key_column` — a unique ID column, if you have one
- `date_columns` — columns with inconsistent date formats
- `currency_columns` — numeric columns stored as text with `$`/commas
- `nonnegative_columns` — columns that should never be negative
- `outlier_columns` — numeric columns to check for extreme outliers
- `canonicalize_columns` — explicit `{"messy value": "canonical value"}` mappings for known
  abbreviations/synonyms
- `auto_canonicalize_columns` — columns to auto-fix for casing/whitespace variants generically
- `titlecase_columns` — text columns to standardize to Title Case
- `email_columns` — columns to validate (and flag, not auto-fix) email format
- `missing_value_strategy` — `{"column": "median" | "mean" | "mode" | "flag_unknown" | "drop_row"}`

Then run:
```bash
python data_cleaning_automation.py your_file.csv --config your_config.json \
    --output-data clean.csv --output-report report.html
```

## Demo results (messy_orders.csv → cleaned_orders.csv)
- 868 → 850 rows (18 duplicates removed)
- 199 missing values fixed
- Completeness: 97.45% → 100%
- 5 region spellings → 5 canonical regions; multiple category spellings → 4 canonical categories
- Mixed date formats → single ISO format; currency-formatted prices → clean numbers
- 99 malformed emails flagged for manual review (not guessed/auto-corrected)
