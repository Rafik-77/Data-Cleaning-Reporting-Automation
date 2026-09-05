#!/usr/bin/env python3
"""
data_cleaning_automation.py
----------------------------
A reusable, configurable engine for cleaning messy tabular data (CSV/Excel)
and auto-generating a data-quality report (HTML, with charts) plus a
cleaned output file.

USAGE (command line):
    python data_cleaning_automation.py <input_file> \
        --output-data cleaned_data.csv \
        --output-report report.html \
        --config config.json     (optional)

USAGE (as a library):
    from data_cleaning_automation import DataCleaner
    cleaner = DataCleaner("messy_orders.csv")
    cleaner.run()
    cleaner.save("cleaned_data.csv")
    cleaner.generate_report("report.html")

WHAT IT HANDLES (generic, works on any dataset with column-type hints or
auto-detection):
    - Duplicate rows (exact duplicates, and duplicate IDs on a key column)
    - Missing values (per-column strategy: mean/median/mode/ffill/drop/flag)
    - Inconsistent text casing & whitespace (trims + standardizes casing)
    - Inconsistent category spellings (fuzzy canonicalization via a mapping
      table you can supply, or auto-grouping of near-identical values)
    - Inconsistent date formats (auto-parses many formats -> one ISO format)
    - Currency-formatted numeric fields ("$1,234.56" -> 1234.56)
    - Invalid values (negative quantities/prices, statistical outliers)
    - Produces a before/after data-quality report with charts

Every step is logged so the report can explain exactly what changed and why.
"""

import argparse
import base64
import io
import json
import re
from collections import Counter
from datetime import datetime

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

plt.rcParams["font.family"] = "DejaVu Sans"

CURRENCY_RE = re.compile(r"[^\d\.\-]")


class DataCleaner:
    """Reusable data cleaning + reporting engine for one tabular dataset."""

    def __init__(self, input_path, config=None):
        self.input_path = input_path
        self.config = config or {}
        self.log = []           # human-readable list of every action taken
        self.metrics = {}        # before/after data-quality metrics
        self.charts = {}         # name -> base64 png
        self.df_raw = None
        self.df = None

    # ---------------------------------------------------------------
    def _log(self, message):
        self.log.append(message)
        print("  -", message)

    # ---------------------------------------------------------------
    def load(self):
        print(f"[1/6] Loading {self.input_path} ...")
        if str(self.input_path).lower().endswith((".xlsx", ".xls")):
            self.df_raw = pd.read_excel(self.input_path)
        else:
            self.df_raw = pd.read_csv(self.input_path)
        self.df = self.df_raw.copy()
        self._log(f"Loaded {len(self.df):,} rows x {len(self.df.columns)} columns")
        return self

    # ---------------------------------------------------------------
    def profile(self, label):
        """Compute a data-quality snapshot; called before and after cleaning."""
        df = self.df
        n_rows = len(df)
        missing_by_col = df.isna().sum()
        total_missing = int(missing_by_col.sum())
        n_dupes = int(df.duplicated().sum())
        dup_key_col = self.config.get("duplicate_key_column")
        n_dupe_keys = int(df.duplicated(subset=[dup_key_col]).sum()) if dup_key_col and dup_key_col in df.columns else None

        completeness = 1 - (total_missing / (n_rows * len(df.columns))) if n_rows else 1

        # Optional: flag structurally invalid values (e.g. malformed emails) --
        # detected and reported, but never auto-"fixed" since guessing the
        # correct value would be unsafe.
        invalid_counts = {}
        email_re = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
        for col in self.config.get("email_columns", []):
            if col in df.columns:
                vals = df[col].dropna().astype(str)
                invalid_counts[col] = int((~vals.str.match(email_re)).sum())

        self.metrics[label] = {
            "rows": n_rows,
            "columns": len(df.columns),
            "missing_total": total_missing,
            "missing_by_column": {c: int(v) for c, v in missing_by_col.items() if v > 0},
            "duplicate_rows": n_dupes,
            "duplicate_keys": n_dupe_keys,
            "completeness_pct": round(completeness * 100, 2),
            "invalid_format_counts": invalid_counts,
        }
        return self.metrics[label]

    # ---------------------------------------------------------------
    def clean(self):
        print("[2/6] Cleaning ...")
        df = self.df
        cfg = self.config

        # --- 1. Standardize column names (strip, no weird spacing) ---
        df.columns = [c.strip() for c in df.columns]

        # --- 2. Remove exact duplicate rows ---
        before = len(df)
        df = df.drop_duplicates().reset_index(drop=True)
        removed = before - len(df)
        if removed:
            self._log(f"Removed {removed} exact duplicate row(s)")

        # --- 3. Remove duplicate keys (keep first), if a key column is configured ---
        dup_key_col = cfg.get("duplicate_key_column")
        if dup_key_col and dup_key_col in df.columns:
            before = len(df)
            df = df.drop_duplicates(subset=[dup_key_col], keep="first").reset_index(drop=True)
            removed_keys = before - len(df)
            if removed_keys:
                self._log(f"Removed {removed_keys} row(s) with duplicate '{dup_key_col}' (kept first occurrence)")

        # --- 4. Trim whitespace on all text/object columns ---
        text_cols = df.select_dtypes(include="object").columns.tolist()
        for col in text_cols:
            trimmed = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
            n_changed = int((trimmed.astype(str) != df[col].astype(str)).sum())
            if n_changed:
                self._log(f"Trimmed stray whitespace in '{col}' ({n_changed} value(s))")
            df[col] = trimmed

        # --- 4b. Standardize casing on name/text columns configured for title-casing
        #          (e.g. "JENNIFER JACKSON" / "jessica thomas" -> "Jennifer Jackson") ---
        titlecase_cols = cfg.get("titlecase_columns", [])
        for col in titlecase_cols:
            if col not in df.columns:
                continue
            new_vals = df[col].apply(lambda x: x.title() if isinstance(x, str) else x)
            n_changed = int((new_vals.astype(str) != df[col].astype(str)).sum())
            if n_changed:
                self._log(f"Standardized capitalization in '{col}' to title case ({n_changed} value(s); "
                          f"note: may mis-case names with unusual capitalization, e.g. 'McDonald')")
            df[col] = new_vals

        # --- 5. Canonicalize categorical columns (casing + known variant mapping) ---
        canonical_cols = cfg.get("canonicalize_columns", {})  # {col: {variant: canonical}}
        for col, mapping in canonical_cols.items():
            if col not in df.columns:
                continue
            def _map(val, mapping=mapping):
                if pd.isna(val):
                    return val
                key = str(val).strip()
                # try exact, then case-insensitive match against mapping keys
                if key in mapping:
                    return mapping[key]
                for variant, canon in mapping.items():
                    if key.lower() == variant.lower():
                        return canon
                return key
            new_vals = df[col].apply(_map)
            n_changed = int((new_vals.astype(str) != df[col].astype(str)).sum())
            if n_changed:
                self._log(f"Standardized {n_changed} inconsistent value(s) in '{col}' "
                          f"(e.g. casing/abbreviation variants -> canonical labels)")
            df[col] = new_vals

        # --- 6. Auto-canonicalize any remaining text columns not explicitly configured
        #         by grouping case/whitespace-insensitive duplicates (e.g. "APPAREL" vs "Apparel") ---
        auto_cols = cfg.get("auto_canonicalize_columns", [])
        for col in auto_cols:
            if col not in df.columns:
                continue
            non_null = df[col].dropna().astype(str)
            groups = {}
            for val in non_null.unique():
                key = re.sub(r"\s+", " ", val.strip()).lower()
                groups.setdefault(key, []).append(val)
            canon_map = {}
            for key, variants in groups.items():
                # canonical form = the most frequent variant, title-cased if all-caps/all-lower
                counts = non_null[non_null.isin(variants)].value_counts()
                canon = counts.idxmax()
                for v in variants:
                    canon_map[v] = canon
            new_vals = df[col].apply(lambda v: canon_map.get(str(v).strip(), v) if pd.notna(v) else v)
            n_changed = int((new_vals.astype(str) != df[col].astype(str)).sum())
            if n_changed:
                self._log(f"Auto-standardized {n_changed} inconsistent spelling/casing variant(s) in '{col}'")
            df[col] = new_vals

        # --- 7. Parse inconsistent date columns into one ISO format ---
        date_cols = cfg.get("date_columns", [])
        for col in date_cols:
            if col not in df.columns:
                continue
            parsed = pd.to_datetime(df[col], errors="coerce", format="mixed")
            n_failed = int(parsed.isna().sum() - df[col].isna().sum())
            if n_failed > 0:
                self._log(f"WARNING: {n_failed} value(s) in '{col}' could not be parsed as dates")
            self._log(f"Standardized '{col}' to ISO date format (yyyy-mm-dd) from mixed formats")
            df[col] = parsed.dt.date

        # --- 8. Clean currency-formatted numeric columns ("$1,234.56" -> 1234.56) ---
        currency_cols = cfg.get("currency_columns", [])
        for col in currency_cols:
            if col not in df.columns:
                continue
            def _to_num(v):
                if pd.isna(v):
                    return np.nan
                if isinstance(v, (int, float)):
                    return float(v)
                cleaned = CURRENCY_RE.sub("", str(v))
                return float(cleaned) if cleaned not in ("", "-", ".") else np.nan
            n_symbols = int(df[col].apply(lambda v: isinstance(v, str) and bool(re.search(r"[$,]", v))).sum())
            df[col] = df[col].apply(_to_num)
            if n_symbols:
                self._log(f"Stripped currency symbols/commas from '{col}' and converted to numeric ({n_symbols} value(s))")

        # --- 9. Fix invalid numeric values (negatives where impossible, extreme outliers) ---
        nonnegative_cols = cfg.get("nonnegative_columns", [])
        for col in nonnegative_cols:
            if col not in df.columns:
                continue
            n_neg = int((df[col] < 0).sum())
            if n_neg:
                df[col] = df[col].abs()
                self._log(f"Corrected {n_neg} negative value(s) in '{col}' (data-entry sign errors -> took absolute value)")

        outlier_cols = cfg.get("outlier_columns", [])
        for col in outlier_cols:
            if col not in df.columns:
                continue
            series = df[col].dropna()
            if len(series) < 10:
                continue
            q1, q3 = series.quantile(0.25), series.quantile(0.75)
            iqr = q3 - q1
            upper = q3 + 3 * iqr
            lower = q1 - 3 * iqr
            mask = (df[col] > upper) | (df[col] < lower)
            n_outliers = int(mask.sum())
            if n_outliers:
                median = series.median()
                df.loc[mask, col] = median
                self._log(f"Replaced {n_outliers} extreme outlier(s) in '{col}' with the column median "
                          f"(IQR method, likely data-entry errors)")

        # --- 10. Handle missing values per configured strategy ---
        missing_strategies = cfg.get("missing_value_strategy", {})
        for col, strategy in missing_strategies.items():
            if col not in df.columns:
                continue
            n_missing = int(df[col].isna().sum())
            if n_missing == 0:
                continue
            if strategy == "median":
                fill = df[col].median()
                df[col] = df[col].fillna(fill)
                self._log(f"Filled {n_missing} missing value(s) in '{col}' with column median ({fill:.2f})")
            elif strategy == "mean":
                fill = df[col].mean()
                df[col] = df[col].fillna(fill)
                self._log(f"Filled {n_missing} missing value(s) in '{col}' with column mean ({fill:.2f})")
            elif strategy == "mode":
                fill = df[col].mode(dropna=True)
                fill = fill.iloc[0] if len(fill) else "Unknown"
                df[col] = df[col].fillna(fill)
                self._log(f"Filled {n_missing} missing value(s) in '{col}' with most common value ('{fill}')")
            elif strategy == "flag_unknown":
                df[col] = df[col].fillna("Unknown")
                self._log(f"Flagged {n_missing} missing value(s) in '{col}' as 'Unknown' (kept row, avoided guessing contact info)")

        # --- Report (but don't silently "fix") structurally invalid values ---
        for col in cfg.get("email_columns", []):
            if col not in df.columns:
                continue
            email_re = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
            vals = df[col].astype(str)
            invalid_mask = (~vals.str.match(email_re)) & df[col].notna() & (df[col] != "Unknown")
            n_invalid = int(invalid_mask.sum())
            if n_invalid:
                self._log(f"Flagged {n_invalid} malformed email(s) in '{col}' for manual review "
                          f"(not auto-corrected — guessing a contact's real address would be unsafe)")
            elif strategy == "drop_row":
                before = len(df)
                df = df.dropna(subset=[col]).reset_index(drop=True)
                self._log(f"Dropped {before - len(df)} row(s) with missing '{col}' (required field)")

        self.df = df
        self._log(f"Cleaning complete: {len(df):,} rows remain")
        return self

    # ---------------------------------------------------------------
    def save(self, output_path):
        print(f"[3/6] Saving cleaned data to {output_path} ...")
        if str(output_path).lower().endswith((".xlsx", ".xls")):
            self.df.to_excel(output_path, index=False)
        else:
            self.df.to_csv(output_path, index=False)
        return self

    # ---------------------------------------------------------------
    def _fig_to_base64(self):
        buf = io.BytesIO()
        plt.savefig(buf, format="png", dpi=140, bbox_inches="tight")
        plt.close()
        buf.seek(0)
        return base64.b64encode(buf.read()).decode()

    def _make_charts(self):
        print("[4/6] Building charts ...")
        before, after = self.metrics["before"], self.metrics["after"]

        # Chart 1: Missing values by column, before vs after
        cols = sorted(set(list(before["missing_by_column"].keys()) + list(after["missing_by_column"].keys())))
        if cols:
            before_vals = [before["missing_by_column"].get(c, 0) for c in cols]
            after_vals = [after["missing_by_column"].get(c, 0) for c in cols]
            x = np.arange(len(cols))
            w = 0.35
            fig, ax = plt.subplots(figsize=(8, 4.2))
            ax.bar(x - w/2, before_vals, w, label="Before", color="#A6202E")
            ax.bar(x + w/2, after_vals, w, label="After", color="#1E7145")
            ax.set_xticks(x)
            ax.set_xticklabels(cols, rotation=25, ha="right", fontsize=8)
            ax.set_title("Missing Values by Column — Before vs. After")
            ax.legend()
            plt.tight_layout()
            self.charts["missing_values"] = self._fig_to_base64()

        # Chart 2: Overall data-quality summary (rows, duplicates, completeness)
        fig, axes = plt.subplots(1, 3, figsize=(12, 3.8))
        labels = ["Before", "After"]
        colors = ["#A6202E", "#1E7145"]

        axes[0].bar(labels, [before["rows"], after["rows"]], color=colors)
        axes[0].set_title("Row Count")
        for i, v in enumerate([before["rows"], after["rows"]]):
            axes[0].text(i, v, f"{v:,}", ha="center", va="bottom", fontsize=9)

        axes[1].bar(labels, [before["duplicate_rows"], after["duplicate_rows"]], color=colors)
        axes[1].set_title("Duplicate Rows")
        for i, v in enumerate([before["duplicate_rows"], after["duplicate_rows"]]):
            axes[1].text(i, v, f"{v:,}", ha="center", va="bottom", fontsize=9)

        axes[2].bar(labels, [before["completeness_pct"], after["completeness_pct"]], color=colors)
        axes[2].set_title("Data Completeness (%)")
        axes[2].set_ylim(0, 105)
        for i, v in enumerate([before["completeness_pct"], after["completeness_pct"]]):
            axes[2].text(i, v, f"{v}%", ha="center", va="bottom", fontsize=9)

        plt.tight_layout()
        self.charts["summary"] = self._fig_to_base64()

    # ---------------------------------------------------------------
    def generate_report(self, output_path):
        print(f"[5/6] Generating report -> {output_path} ...")
        self._make_charts()
        before, after = self.metrics["before"], self.metrics["after"]
        now = datetime.now().strftime("%Y-%m-%d %H:%M")

        log_html = "".join(f"<li>{entry}</li>" for entry in self.log)
        rows_removed = before["rows"] - after["rows"]
        invalid_emails = sum(after.get("invalid_format_counts", {}).values())

        html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Data Cleaning & Quality Report</title>
<style>
  body {{ font-family: Arial, Helvetica, sans-serif; background:#F4F5F7; color:#222; margin:0; padding:0; }}
  .wrap {{ max-width: 980px; margin: 0 auto; padding: 32px 24px 60px; }}
  h1 {{ color:#1F2A44; margin-bottom:4px;}}
  .subtitle {{ color:#666; margin-top:0; margin-bottom:28px; }}
  .kpi-row {{ display:flex; gap:16px; flex-wrap:wrap; margin-bottom:28px; }}
  .kpi {{ background:#fff; border-radius:10px; padding:16px 20px; flex:1; min-width:160px;
          box-shadow:0 1px 3px rgba(0,0,0,0.08); }}
  .kpi .label {{ font-size:12px; color:#666; text-transform:uppercase; letter-spacing:.03em;}}
  .kpi .value {{ font-size:26px; font-weight:700; color:#1E7145; margin-top:4px;}}
  .kpi .value.neutral {{ color:#2F5597; }}
  section {{ background:#fff; border-radius:10px; padding:24px; margin-bottom:24px;
             box-shadow:0 1px 3px rgba(0,0,0,0.08); }}
  section h2 {{ margin-top:0; color:#1F2A44; font-size:18px; border-bottom:2px solid #F0F0F0; padding-bottom:10px;}}
  img {{ max-width:100%; border-radius:6px; }}
  table {{ border-collapse:collapse; width:100%; font-size:13px; }}
  th, td {{ text-align:left; padding:8px 10px; border-bottom:1px solid #eee; }}
  th {{ background:#1F2A44; color:#fff; }}
  ul.log {{ font-size:13px; line-height:1.9; padding-left:20px; }}
  .footer {{ text-align:center; color:#999; font-size:12px; margin-top:20px; }}
  .badge {{ display:inline-block; padding:2px 10px; border-radius:12px; font-size:12px; font-weight:600; }}
  .badge.good {{ background:#E2F0D9; color:#1E7145; }}
</style>
</head>
<body>
<div class="wrap">
  <h1>Data Cleaning &amp; Quality Report</h1>
  <p class="subtitle">Source file: <b>{self.input_path}</b> &nbsp;|&nbsp; Generated {now}
     &nbsp;|&nbsp; <span class="badge good">Automated pipeline</span></p>

  <div class="kpi-row">
    <div class="kpi"><div class="label">Rows (before → after)</div>
        <div class="value neutral">{before['rows']:,} → {after['rows']:,}</div></div>
    <div class="kpi"><div class="label">Rows Removed</div>
        <div class="value">{rows_removed:,}</div></div>
    <div class="kpi"><div class="label">Missing Values Fixed</div>
        <div class="value">{before['missing_total'] - after['missing_total']:,}</div></div>
    <div class="kpi"><div class="label">Completeness</div>
        <div class="value">{before['completeness_pct']}% → {after['completeness_pct']}%</div></div>
    <div class="kpi"><div class="label">Flagged for Manual Review</div>
        <div class="value" style="color:#C55A11;">{invalid_emails:,}</div></div>
  </div>

  <section>
    <h2>Data Quality: Before vs. After</h2>
    <img src="data:image/png;base64,{self.charts.get('summary','')}" alt="Summary chart">
  </section>

  {"<section><h2>Missing Values by Column</h2><img src='data:image/png;base64," + self.charts.get('missing_values','') + "' alt='Missing values chart'></section>" if 'missing_values' in self.charts else ""}

  <section>
    <h2>Cleaning Actions Log</h2>
    <ul class="log">{log_html}</ul>
  </section>

  <section>
    <h2>Cleaned Data Preview (first 10 rows)</h2>
    {self.df.head(10).to_html(index=False, border=0)}
  </section>

  <div class="footer">Generated automatically by data_cleaning_automation.py — re-run on any new export to refresh this report.</div>
</div>
</body>
</html>"""

        with open(output_path, "w") as f:
            f.write(html)
        return self

    # ---------------------------------------------------------------
    def run(self):
        self.load()
        self.profile("before")
        self.clean()
        self.profile("after")
        print("[6/6] Done.")
        return self


# =====================================================================
# Default configuration for the demo "messy_orders.csv" dataset.
# For a NEW dataset, write your own config dict/JSON describing which
# columns need which treatment -- everything else in the engine above
# is fully generic and dataset-agnostic.
# =====================================================================
DEMO_CONFIG = {
    "duplicate_key_column": "OrderID",
    "date_columns": ["OrderDate"],
    "currency_columns": ["UnitPrice"],
    "nonnegative_columns": ["Quantity"],
    "outlier_columns": ["UnitPrice"],
    # Explicit mapping handles true abbreviations/synonyms (not just casing/whitespace,
    # which auto_canonicalize_columns below handles generically for anything left over).
    "canonicalize_columns": {
        "Region": {"N.": "North", "S.": "South", "E.": "East", "W.": "West", "Ctrl": "Central"},
        "Category": {"Clothing": "Apparel", "Electronic": "Electronics",
                      "Home and Living": "Home & Living", "Home&Living": "Home & Living"},
    },
    "auto_canonicalize_columns": ["Region", "Category"],
    "email_columns": ["Email"],
    "titlecase_columns": ["CustomerName"],
    "missing_value_strategy": {
        "Email": "flag_unknown",
        "Phone": "flag_unknown",
        "Quantity": "median",
        "UnitPrice": "median",
        "Region": "mode",
    },
}


def main():
    parser = argparse.ArgumentParser(description="Clean a messy CSV/Excel file and generate a data-quality report.")
    parser.add_argument("input_file", help="Path to the raw CSV or Excel file")
    parser.add_argument("--output-data", default="cleaned_data.csv", help="Path for the cleaned output file")
    parser.add_argument("--output-report", default="report.html", help="Path for the generated HTML report")
    parser.add_argument("--config", default=None, help="Path to a JSON config file (see DEMO_CONFIG for the shape)")
    args = parser.parse_args()

    if args.config:
        with open(args.config) as f:
            config = json.load(f)
    else:
        config = DEMO_CONFIG
        print("(No --config given: using the built-in demo config. Pass your own JSON config for a different schema.)")

    cleaner = DataCleaner(args.input_file, config=config)
    cleaner.run()
    cleaner.save(args.output_data)
    cleaner.generate_report(args.output_report)
    print(f"\nCleaned data:  {args.output_data}")
    print(f"Report:        {args.output_report}")


if __name__ == "__main__":
    main()
