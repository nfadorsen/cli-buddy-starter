# Copilot Instructions (sample)

Drop this file into your own repo at `.github/copilot-instructions.md` and edit.
It captures enterprise-software defaults for working on a Windows machine where
Office files are protected by sensitivity labels (IRM). The skills in this repo
work without this file — it just makes the assistant's behavior more predictable.

---

## Core behaviors

- **Never claim you cannot do something without first checking.** Investigate before declining.
- **Be direct about capabilities and limitations** — don't default to "I can't."
- **Command approvals:** before running a command or asking me to approve it, explain in plain English what it does, whether it changes files, whether it uses the network, and whether it's reversible.
- **Risk label:** for any command that needs execution or approval, assign a clear risk label (`Low`, `Medium`, or `High`) with a one-line reason.

## SharePoint & file access

- Microsoft SharePoint URLs cannot be fetched directly. If you need a deeper read of a file, ask me to save a local copy and work from the local file.
- For Excel files, default to the `excel-toolkit` skill. If the workbook is IRM/sensitivity-labeled (OLE2 on disk) and Python libraries can't read it, use `excel-enterprise`. Never remove or downgrade sensitivity labels unless I explicitly confirm it's permitted.

## Default labeling assumption

- Office files in this tenant often default to a sensitivity label that produces IRM-protected / OLE2-wrapped files on disk even after "Save As".
- Python parsers (`openpyxl`, `python-pptx`, `markitdown`, etc.) frequently fail on these with `BadZipFile` or silently parse nothing.
- **Detect first, then choose lane.** Sniff the first bytes of any Office file (`.xlsx`, `.pptx`, `.docx`):
  - `50 4B` (PK) → zip-backed, Python is safe.
  - `D0 CF 11 E0` → OLE2 / IRM, use the enterprise wrapper skill (`excel-enterprise` / `pptx-enterprise` / `docx-enterprise`) which routes to Excel / PowerPoint / Word COM.
- PDFs are different. Password-encrypted PDFs are handled by `pypdf` / `qpdf`; Azure-RMS-wrapped PDFs (rare) should be opened in Adobe Reader and saved unprotected rather than chased with Python.
- Never remove or downgrade sensitivity labels without explicit permission.
- **Do not chase Python workarounds for IRM-encrypted Office files.** This is architectural:
  - MIP SDK has no Python bindings.
  - Graph API does not return decrypted content for IRM-protected files.
  - Desktop COM is the only durable local path — route straight to the enterprise skill when the header sniffs as `D0 CF 11 E0`.

## Office COM hygiene

- Every COM session that touches Excel can leave a headless `EXCEL.EXE` orphan. Sweep these BEFORE starting a new session (kill processes with an empty `MainWindowTitle` by specific PID) and track new Office PIDs so any you spawned can be force-killed on exit. Never kill visible Office windows — those are the user's.
- PowerShell → COM marshaling is strict. Cast `Cells.Item` row/column arguments to `[int]` and numeric cell values to `[double]`. Bare integer arithmetic like `$i + 2` produces a `Double` and throws a cast error against COM setters.
- For IRM/OLE2 decks, use `Presentation.Save()` to write in place. Do NOT use `SaveAs` with `ppSaveAsOpenXMLPresentation=24` — it can re-wrap or mutate the IRM container and strip or downgrade the sensitivity label.
- For rotated or complex shapes (braces, callouts with non-zero `Rotation`), do NOT resize or reposition via direct property mutation. Rotation math does not compose cleanly through COM. Instead, `Copy()` the known-good shape and `Paste()` onto the target slide.

## Waterfall charts (ChartType=119)

- Waterfall chart **structure** is effectively read-only through COM: `Chart.SetSourceData`, `Series.XValues`, `Series.Values`, `Series.Formula`, and `Point.IsTotal` all throw `Unexpected HRESULT`.
- What DOES work: writing values into the embedded workbook cells (`Chart.ChartData.Workbook.Worksheets(1).Cells(...)`) and persisting via `Presentation.Save()`.
- Durable pattern:
  1. Duplicate the source slide so the new slide inherits chart formatting and subtotal flags.
  2. Overwrite embedded workbook cells with new category labels and values.
  3. `Workbook.Close($true)` then `Presentation.Save()`.
  4. Tell the user up-front that two things require a manual UI fix: (a) range extension if row count changed, (b) "Set as Total" toggles if subtotal position changed. Budget ~30 seconds of UI work instead of hours fighting COM.

## Building internal / enterprise software

### Core philosophy
- Optimize for simplicity, clarity, and safety over cleverness.
- Prefer fewer reliable features over advanced features that add complexity.
- Assume end users are non-technical unless stated otherwise.
- Prefer transparent, inspectable scripts over binaries or hidden behavior.

### Security & privacy (non-negotiable)
- No silent background execution, scheduled tasks, services, or persistence mechanisms unless explicitly approved.
- No registry writes (especially HKLM), no admin / elevation without approval.
- No telemetry. No data leaves the machine unless explicitly documented.
- Any diagnostic artifact described as "sanitized" MUST actually be sanitized (no usernames, emails, tenant IDs, tokens, full local paths).

### Logging & diagnostics
- Plain text, local-only, timestamped, step-named.
- Never contains sensitive data.
- Safe to attach in email or Teams.

### Failure modes
- Fail fast, fail clearly, fail safely.
- Never leave the system in a half-updated or ambiguous state.
- When in doubt, stop and instruct the user to re-run the primary entry point.

---

Feel free to add project-specific sections below (priorities, key people, domain conventions). Keep anything sensitive OUT of this file if the repo is public.
