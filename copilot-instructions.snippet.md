<!-- BEGIN: cli-buddy-starter enterprise skills v1 -->
## Enterprise Office skills (IRM-aware)

This block is an **appendable snippet** that pairs with the `pptx-enterprise`,
`docx-enterprise`, and `excel-enterprise` skills from
https://github.com/nfadorsen/cli-buddy-starter. Paste it at the end of your
existing `.github/copilot-instructions.md` — do NOT replace anything above.
Keep the `BEGIN` / `END` markers so it's easy to update or remove later.

### Detect-first for Office files

Before opening any `.xlsx`, `.pptx`, or `.docx` file with a Python library,
sniff the first 4 bytes of the file:

- `50 4B 03 04` (PK) → zip-backed OpenXML. Python tools are safe.
- `D0 CF 11 E0` → OLE2 compound document. Could be a legacy `.ppt` / `.xls` OR
  an IRM / sensitivity-labeled OOXML file. **Python parsers will fail.** Route
  straight to the matching enterprise skill (`pptx-enterprise` /
  `docx-enterprise` / `excel-enterprise`), which uses Office COM.

Each enterprise skill has a `detect_container.ps1` helper that does this sniff
and returns JSON. Use it — don't re-implement the check.

### Sensitivity labels (non-negotiable)

- **Never remove or downgrade a sensitivity / IRM label** without my explicit
  permission. If a Python library can't read a labeled file, use the COM path
  in the enterprise skill instead of asking me to strip the label.
- For IRM-wrapped Office files, prefer `Document.Save()` / `Presentation.Save()`
  / `Workbook.Save()` to write in place. **Avoid `SaveAs`** — it can re-wrap or
  mutate the IRM container and strip the label.
- Never edit the source file in place on an IRM-wrapped document unless I
  explicitly ask for it. Default to writing a new file or a clean copy.

### Office COM hygiene

- Every COM session that touches Excel can leave a headless `EXCEL.EXE`
  orphan behind (no `MainWindowTitle`). Sweep these **before** starting a new
  session by specific PID. Never kill a visible Office window — that's mine.
- Track which Office PIDs existed before the session started. On cleanup,
  kill only the PIDs the session itself spawned.
- PowerShell → COM marshaling is strict. Cast `Cells.Item` row/column args to
  `[int]` and numeric cell values to `[double]`. Bare integer arithmetic
  produces a `Double` and throws a cast error against COM setters.
- If Excel COM fails with `0x80080005 CO_E_SERVER_EXEC_FAILURE`, abort and
  surface the remediation (close Excel, kill orphans, open Excel once
  interactively). Do NOT fall back to screenshots or shape-drawn "fake" charts.

### Waterfall charts (ChartType=119)

If the task edits a waterfall chart inside an IRM-wrapped deck:

- Chart **structure** is effectively read-only via COM: `SetSourceData`,
  `Series.XValues`, `Series.Values`, `Series.Formula`, and `Point.IsTotal` all
  throw `Unexpected HRESULT`. Don't fight it.
- What DOES work: writing values into the embedded workbook cells
  (`Chart.ChartData.Workbook.Worksheets(1).Cells(...)`) and persisting via
  `Presentation.Save()`.
- Durable pattern: duplicate the source slide → overwrite the embedded
  workbook cells → `Presentation.Save()`. Tell me up-front that two things
  need a manual UI fix: (a) range extension if row count changed,
  (b) "Set as Total" toggles if the subtotal position changed.

### Output location

Write all skill outputs under an `exports\` folder next to the source file.
If that location isn't writable, fall back to `.\exports\` in the current
working directory. Never write outputs somewhere else silently.

<!-- END: cli-buddy-starter enterprise skills v1 -->
