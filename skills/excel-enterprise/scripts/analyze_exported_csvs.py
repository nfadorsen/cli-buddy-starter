#!/usr/bin/env python3
"""
Analyze CSVs exported by export_labeled_excel.ps1.

Reads <exports-dir>/sheets.json, loads each referenced CSV with pandas, and
emits:
  - <exports-dir>/analysis_report.json   (machine-readable)
  - a markdown summary to stdout         (paste into chat)

Heuristics:
  * Missing values: count + % per column; top offenders per sheet.
  * ID-like columns: header matches /(id|order|po|ro|msft|vendor|device|sku|
    asin|gtin|ean|upc)/i AND column is numeric integer-like -> flag as
    "store as string" (leading-zero risk).
  * Date-like strings: object dtype AND pd.to_datetime(..., errors='coerce')
    parses >= 80% of non-null values -> flag.
  * Mixed-type columns: object dtype AND a non-trivial mix of numeric-parseable
    and non-numeric-text values.
  * Top 5 insights: combines largest missing %, largest numeric totals on
    obvious amount columns, and flagged data-quality issues.

Usage:
  python analyze_exported_csvs.py --exports-dir <DIR>
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

try:
    import pandas as pd
except Exception as e:  # pragma: no cover
    print(f"ERROR: pandas is required. Install with: python -m pip install pandas\n{e}", file=sys.stderr)
    sys.exit(2)

ID_HEADER_RE = re.compile(
    r"(?<![a-z])(id|order|po|ro|msft|vendor|device|sku|asin|gtin|ean|upc)(?![a-z])",
    re.IGNORECASE,
)
AMOUNT_HEADER_RE = re.compile(
    r"(amount|amt|value|revenue|cost|spend|savings|total|qty|quantity|units|usd|dollars|price)",
    re.IGNORECASE,
)


def load_manifest(exports_dir: Path) -> list[dict]:
    manifest_path = exports_dir / "sheets.json"
    if not manifest_path.exists():
        raise FileNotFoundError(f"sheets.json not found in {exports_dir}. Run export_labeled_excel.ps1 first.")
    with manifest_path.open("r", encoding="utf-8-sig") as fh:
        return json.load(fh)


def looks_numeric(series: pd.Series) -> tuple[int, int]:
    """Return (parseable_count, non_null_count) for object series."""
    non_null = series.dropna().astype(str).str.strip()
    non_null = non_null[non_null != ""]
    if non_null.empty:
        return 0, 0
    parsed = pd.to_numeric(non_null, errors="coerce")
    return int(parsed.notna().sum()), int(len(non_null))


def looks_datelike(series: pd.Series) -> tuple[int, int]:
    non_null = series.dropna().astype(str).str.strip()
    non_null = non_null[non_null != ""]
    if non_null.empty:
        return 0, 0
    # Suppress pandas parsing warnings by catching errors='coerce' behavior.
    parsed = pd.to_datetime(non_null, errors="coerce", utc=False)
    return int(parsed.notna().sum()), int(len(non_null))


def is_integer_numeric(series: pd.Series) -> bool:
    """True if numeric column and all non-null values are integers (no fractional part)."""
    if not pd.api.types.is_numeric_dtype(series):
        return False
    s = series.dropna()
    if s.empty:
        return False
    try:
        return bool((s % 1 == 0).all())
    except Exception:
        return False


def analyze_sheet(sheet_meta: dict) -> dict:
    name = sheet_meta.get("name", "<unknown>")
    csv_path = Path(sheet_meta["csv"])
    report: dict[str, Any] = {
        "name": name,
        "csv": str(csv_path),
        "rows_reported_by_excel": sheet_meta.get("rows"),
        "cols_reported_by_excel": sheet_meta.get("cols"),
        "error": None,
    }
    if not csv_path.exists():
        report["error"] = f"CSV missing: {csv_path}"
        return report

    # Read as strings first so we can detect mixed/date-like patterns honestly,
    # then coerce a numeric view for stats.
    try:
        df_raw = pd.read_csv(csv_path, dtype=str, keep_default_na=True, na_values=["", "NA", "NaN", "null"])
    except Exception as e:
        report["error"] = f"Failed to read CSV: {e}"
        return report

    report["rows_csv"] = int(df_raw.shape[0])
    report["cols_csv"] = int(df_raw.shape[1])

    missing = []
    suspicious_id = []
    suspicious_date = []
    suspicious_mixed = []
    numeric_totals: list[tuple[str, float, int]] = []  # (col, sum, n)

    total_rows = max(df_raw.shape[0], 1)

    for col in df_raw.columns:
        s_raw = df_raw[col]
        miss = int(s_raw.isna().sum() + (s_raw.fillna("").astype(str).str.strip() == "").sum() - s_raw.isna().sum())
        # Simpler: any blank-after-strip
        miss = int((s_raw.fillna("").astype(str).str.strip() == "").sum())
        pct = round(100.0 * miss / total_rows, 1)
        missing.append({"column": col, "missing": miss, "pct": pct})

        # Numeric-ness
        num_ok, nn = looks_numeric(s_raw)
        is_numeric_col = nn > 0 and num_ok / nn >= 0.9

        # Date-ness (only consider if not already numeric)
        if not is_numeric_col:
            dt_ok, dt_nn = looks_datelike(s_raw)
            if dt_nn > 0 and dt_ok / dt_nn >= 0.8:
                suspicious_date.append({
                    "column": col,
                    "parsed_pct": round(100.0 * dt_ok / dt_nn, 1),
                    "sample_nonnull": int(dt_nn),
                })

        # Mixed-type (object column with both numeric-parseable AND non-numeric text)
        if nn >= 10:
            non_numeric = nn - num_ok
            if 0.1 <= (num_ok / nn) <= 0.9 and num_ok >= 3 and non_numeric >= 3:
                suspicious_mixed.append({
                    "column": col,
                    "numeric_pct": round(100.0 * num_ok / nn, 1),
                    "sample_nonnull": int(nn),
                })

        # ID-like
        if ID_HEADER_RE.search(str(col) or ""):
            if is_numeric_col:
                s_num = pd.to_numeric(s_raw, errors="coerce")
                if is_integer_numeric(s_num):
                    suspicious_id.append({
                        "column": col,
                        "reason": "Header looks like an ID and column is integer-numeric; store as string to preserve leading zeros and exact values.",
                        "sample_nonnull": int(nn),
                    })

        # Numeric totals on amount-ish columns
        if AMOUNT_HEADER_RE.search(str(col) or "") and is_numeric_col:
            s_num = pd.to_numeric(s_raw, errors="coerce")
            numeric_totals.append((col, float(s_num.sum(skipna=True)), int(s_num.notna().sum())))

    missing_sorted = sorted(missing, key=lambda r: r["missing"], reverse=True)
    top_missing = [m for m in missing_sorted if m["missing"] > 0][:10]

    # Build insights
    insights: list[str] = []
    # 1) Sheet size
    insights.append(
        f"Sheet '{name}' has {report['rows_csv']} rows x {report['cols_csv']} cols after CSV export."
    )
    # 2) Biggest numeric total, if any
    numeric_totals.sort(key=lambda t: abs(t[1]), reverse=True)
    if numeric_totals:
        col, total, n = numeric_totals[0]
        insights.append(f"Largest amount-like column: '{col}' sums to {total:,.0f} across {n} values.")
    # 3) Highest missing %
    if top_missing:
        t = top_missing[0]
        insights.append(f"Most-missing column: '{t['column']}' ({t['missing']} blanks, {t['pct']}%).")
    # 4) ID-as-number warning
    if suspicious_id:
        ids = ", ".join(f"'{x['column']}'" for x in suspicious_id[:3])
        insights.append(f"ID-like columns stored as numbers: {ids}. Convert to string to avoid leading-zero loss.")
    # 5) Dates-as-text warning
    if suspicious_date:
        ds = ", ".join(f"'{x['column']}'" for x in suspicious_date[:3])
        insights.append(f"Date-like columns stored as text: {ds}. Parse to datetime before use.")
    # Pad to exactly 5 if we have room
    extras = []
    if suspicious_mixed:
        ms = ", ".join(f"'{x['column']}'" for x in suspicious_mixed[:3])
        extras.append(f"Mixed-type columns (numbers and text together): {ms}.")
    if len(top_missing) > 1:
        extras.append(
            "Multiple columns have missing data; consider a data-quality review before distribution."
        )
    for e in extras:
        if len(insights) >= 5:
            break
        insights.append(e)
    insights = insights[:5]

    report["missing_top10"] = top_missing
    report["suspicious_id_columns"] = suspicious_id
    report["suspicious_date_columns"] = suspicious_date
    report["suspicious_mixed_columns"] = suspicious_mixed
    report["numeric_totals_top"] = [
        {"column": c, "sum": t, "n": n} for c, t, n in numeric_totals[:5]
    ]
    report["insights"] = insights
    return report


def render_markdown(overall: dict) -> str:
    lines: list[str] = []
    lines.append(f"## Analysis of `{overall['workbook']}`")
    lines.append("")
    lines.append(f"Sheets: **{len(overall['sheets'])}**  |  Exports dir: `{overall['exports_dir']}`")
    lines.append("")
    for sh in overall["sheets"]:
        lines.append(f"### Sheet: `{sh['name']}`")
        if sh.get("error"):
            lines.append(f"- ERROR: {sh['error']}")
            lines.append("")
            continue
        lines.append(f"- Rows x Cols (CSV): **{sh['rows_csv']} x {sh['cols_csv']}**")
        if sh.get("missing_top10"):
            lines.append("- Top missing columns:")
            for m in sh["missing_top10"][:5]:
                lines.append(f"  - `{m['column']}` - {m['missing']} blanks ({m['pct']}%)")
        if sh.get("suspicious_id_columns"):
            lines.append("- ID-like columns stored as numbers (store as string):")
            for x in sh["suspicious_id_columns"]:
                lines.append(f"  - `{x['column']}`")
        if sh.get("suspicious_date_columns"):
            lines.append("- Date-like columns stored as text:")
            for x in sh["suspicious_date_columns"]:
                lines.append(f"  - `{x['column']}` ({x['parsed_pct']}% parseable)")
        if sh.get("suspicious_mixed_columns"):
            lines.append("- Mixed-type object columns:")
            for x in sh["suspicious_mixed_columns"]:
                lines.append(f"  - `{x['column']}` ({x['numeric_pct']}% numeric)")
        if sh.get("insights"):
            lines.append("- Insights:")
            for ins in sh["insights"]:
                lines.append(f"  1. {ins}")
        lines.append("")
    return "\n".join(lines)


def main() -> int:
    ap = argparse.ArgumentParser(description="Analyze CSVs exported by export_labeled_excel.ps1")
    ap.add_argument("--exports-dir", required=True, help="Directory containing sheets.json and per-sheet CSVs.")
    args = ap.parse_args()

    exports_dir = Path(args.exports_dir).resolve()
    if not exports_dir.exists():
        print(f"ERROR: exports dir does not exist: {exports_dir}", file=sys.stderr)
        return 2

    manifest = load_manifest(exports_dir)
    overall = {
        "workbook": str(exports_dir.parent),  # best-effort
        "exports_dir": str(exports_dir),
        "sheets": [analyze_sheet(m) for m in manifest],
    }

    report_path = exports_dir / "analysis_report.json"
    with report_path.open("w", encoding="utf-8") as fh:
        json.dump(overall, fh, indent=2, default=str)

    print(render_markdown(overall))
    print(f"\n_Report JSON: {report_path}_")
    return 0


if __name__ == "__main__":
    sys.exit(main())
