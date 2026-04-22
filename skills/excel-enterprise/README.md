# excel-enterprise (Copilot CLI skill)

## Why this skill exists

Microsoft enterprise workbooks are frequently protected with a **sensitivity
label** (Microsoft Information Protection / IRM). On disk, a labeled `.xlsx`
is not a normal zip file — it's an encrypted OLE2 compound container. Python
libraries like `openpyxl` and `pandas.read_excel` will fail on these with
`zipfile.BadZipFile: File is not a zip file`.

This skill wraps the standard `excel-toolkit` skill so that:

1. **Normal .xlsx/.xlsm** — still fast, handled by Python directly.
2. **IRM-wrapped workbooks** — the skill automatically switches to Excel COM,
   exports each worksheet to CSV **without modifying the source or the label**,
   and then analyzes the CSVs with pandas.

## What it never does

- Never removes or downgrades a sensitivity label.
- Never writes to the source workbook.
- Never sends data off the machine.

## Layout

```
excel-enterprise\
  SKILL.md
  README.md                          (this file)
  scripts\
    export_labeled_excel.ps1         (COM-based CSV exporter)
    analyze_exported_csvs.py         (pandas analyzer)
```

Outputs always go to an `exports\` folder next to the workbook (or `cwd\exports`
if the workbook folder is read-only).

## How to test

### Case A — normal .xlsx (Python path)

1. Pick any unprotected `.xlsx` (e.g., something you created yourself and saved
   with no sensitivity label).
2. In Copilot CLI, ask:
   > "Inspect `<path to normal .xlsx>` — sheet list, row/col counts, missing
   > values, and top 5 insights."
3. The skill should call `excel-toolkit`'s `inspect_excel.py` and
   `analyze_excel.py` and return a summary. No `exports\` folder is created.

### Case B — IRM-wrapped .xlsx (COM export path)

1. Pick a workbook that has a Microsoft sensitivity label applied (Confidential,
   Highly Confidential, etc.).
2. Verify it's OLE2-wrapped by checking the first 4 bytes in PowerShell:
   ```powershell
   $fs = [IO.File]::Open('<PATH>','Open','Read','ReadWrite')
   $b = New-Object byte[] 4; $fs.Read($b,0,4) | Out-Null; $fs.Close()
   ($b | % { $_.ToString('X2') }) -join ' '
   ```
   You should see `D0 CF 11 E0`.
3. In Copilot CLI:
   > "Analyze `<path to labeled .xlsx>`. Use the COM fallback, export sheets
   > to CSV, and summarize."
4. The skill should:
   - Run `scripts\export_labeled_excel.ps1 -Path <PATH>`.
   - Create `exports\<sheet>.csv` for each worksheet plus `exports\sheets.json`.
   - Run `scripts\analyze_exported_csvs.py --exports-dir <exports-dir>`.
   - Produce `exports\analysis_report.json` and print a markdown summary.

### Troubleshooting

- **"Server execution failed" on first COM call** — Excel was starting cold.
  The exporter retries up to 5 times; if it still fails, close any lingering
  `EXCEL.EXE` in Task Manager and re-run.
- **"Update links" or "Enable editing" dialog in Excel** — choose
  *Don't update* and *Enable editing* (the workbook opens ReadOnly anyway),
  then re-run.
- **Sheet name with illegal filename characters** — the exporter sanitizes
  `/ \ : * ? " < > |` (and a few more) to `_`. If two sheets collapse to the
  same safe name, the exporter appends `_2`, `_3`, etc.
