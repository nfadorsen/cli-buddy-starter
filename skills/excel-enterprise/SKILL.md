---
name: excel-enterprise
description: "Enterprise Excel workflow. Uses excel-toolkit for normal .xlsx/.csv. Detects IRM/sensitivity-labeled OLE2-wrapped workbooks and falls back to Excel COM export, then analyzes exported CSVs."
---

# Excel Enterprise (IRM-aware wrapper)

Use this skill **instead of `excel-toolkit` alone** when working with workbooks on
an enterprise Windows machine where files may be protected by Microsoft
Information Protection (sensitivity labels / IRM). The skill prefers the fast
Python path and automatically falls back to Excel COM when — and only when — the
file is IRM-wrapped.

This skill never modifies the source workbook and never changes sensitivity
labels.

---

## When to use this skill

Trigger whenever the user asks to **inspect, read, or analyze** an `.xlsx`,
`.xlsm`, or `.xls` file on this machine. For pure CSV/TSV files, defer to
`excel-toolkit` directly.

---

## Step 1 — Detect file type (magic-byte sniff) — MANDATORY FIRST STEP

**Always detect before attempting Python.** In this tenant, Office files default to the "Confidential \ Internal Only" label and are frequently OLE2/IRM-wrapped on disk — `openpyxl` and `pandas.read_excel` will fail on these with `BadZipFile`. Detect-first avoids the wasted round-trip.

Use the bundled script (works even if Excel has the file open):

```powershell
pwsh -File "$env:USERPROFILE\.copilot\skills\excel-enterprise\scripts\detect_container.ps1" -Path "<PATH>"
```

Returns JSON `{ path, headerHex, containerType }` where `containerType` is `zip`, `ole2`, or `unknown`. Route:

- `zip` → **Python path (Step 2a)**
- `ole2` or `unknown` → **COM fallback path (Step 2b)** — do NOT attempt Python first

Fallback one-liner (if the script isn't available):

```powershell
$fs = [System.IO.File]::Open('<PATH>','Open','Read','ReadWrite')
$b = New-Object byte[] 4; $fs.Read($b,0,4) | Out-Null; $fs.Close()
($b | ForEach-Object { $_.ToString('X2') }) -join ' '
```

Interpretation:

| Magic bytes     | Meaning                                                | Route       |
|-----------------|--------------------------------------------------------|-------------|
| `50 4B 03 04`   | Zip-backed OpenXML (normal .xlsx / .xlsm)              | Python      |
| `D0 CF 11 E0`   | OLE2 compound — legacy .xls OR IRM-wrapped OOXML       | COM fallback|
| other           | Unknown — try Python; on `BadZipFile`, use COM fallback|             |

If `openpyxl` or `pandas.read_excel` raises `zipfile.BadZipFile: File is not a
zip file`, treat the workbook as IRM-wrapped and switch to the COM fallback —
do not retry in Python.

---

## Step 2a — Python path (normal files)

Delegate to the existing `excel-toolkit` skill. From `excel-enterprise`, invoke:

```powershell
python "$env:USERPROFILE\.copilot\skills\excel-toolkit\scripts\setup_deps.py"
python "$env:USERPROFILE\.copilot\skills\excel-toolkit\scripts\inspect_excel.py" "<PATH>" --data --rows 20
python "$env:USERPROFILE\.copilot\skills\excel-toolkit\scripts\analyze_excel.py" "<PATH>" --correlations
```

Summarize the results to the user as usual.

---

## Step 2b — COM fallback path (IRM-wrapped files)

Two steps: export, then analyze.

### Step 2b.1 — Export each sheet to CSV

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File `
  "$env:USERPROFILE\.copilot\skills\excel-enterprise\scripts\export_labeled_excel.ps1" `
  -Path "<PATH>"
```

Optional `-OutDir "<DIR>"` to override the default output location.

Default `-OutDir` resolution: `<workbook-folder>\exports`. If that folder is
not writable, the script falls back to `<cwd>\exports`.

Outputs:
- `exports\<sheet-name>.csv` — one CSV per worksheet (sheet name sanitized)
- `exports\sheets.json` — `[{ name, rows, cols, csv }]`

The script opens the workbook **ReadOnly** and copies each sheet to a fresh
temporary workbook before saving CSV, so the source workbook is never mutated.

### Step 2b.2 — Analyze the exported CSVs

```powershell
python "$env:USERPROFILE\.copilot\skills\excel-enterprise\scripts\analyze_exported_csvs.py" `
  --exports-dir "<workbook-folder>\exports"
```

Outputs:
- `exports\analysis_report.json` — machine-readable report (sheets, missing
  values, suspicious columns, insights)
- Markdown summary printed to **stdout** — paste this into the chat response.

---

## Step 3 — Present findings

Always include, per sheet:

1. Row × column counts.
2. Top missing-value columns (count and %).
3. Flagged suspicious columns:
   - **ID-like numeric columns** that should be strings (leading-zero risk).
   - **Date-like text columns** (dates stored as strings).
   - **Mixed-type object columns** (numbers and text in the same column).
4. Top 5 insights (largest numeric totals, concentrations, outliers,
   data-quality issues likely to bite a downstream consumer).

Keep language plain and business-focused.

---

## Safety rails (do not bypass)

- **Never modify the source workbook.** Use ReadOnly opens and copy-to-temp
  before CSV export.
- **Never remove, downgrade, or suggest removing a sensitivity label** to make
  Python happy. If the user is blocked, explain the options (CSV export via
  this skill, or label adjustment by the user themselves) and stop.
- **Write outputs only to `exports\`** beside the workbook, or `<cwd>\exports`
  if the workbook folder is not writable. Never write anywhere else.
- **No network calls.** Everything runs locally.
- **Excel COM prompts:** if Excel shows an update-links, protected-view, or
  enable-content prompt during export, briefly tell the user what to click
  (typically "Don't update" and "Enable editing"), then re-run the exporter.
  Do not automate those clicks.

---

## Example prompts

**Normal .xlsx (Python route):**
> "Inspect `C:\Users\<you>\Downloads\Sales.xlsx`, show sheet counts,
> missing values, and top 5 insights."

**IRM-wrapped .xlsx (COM export route):**
> "Analyze `C:\path\to\Workbook.xlsx`. It's sensitivity-labeled so Python
> can't open it — use the COM fallback, export the sheets to CSV, and give
> me the standard summary."
