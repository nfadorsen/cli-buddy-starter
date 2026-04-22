---
name: pptx-enterprise
description: "Enterprise PowerPoint workflow. Reads, analyzes, edits, and creates decks. Detects legacy/IRM-wrapped OLE2 containers (D0 CF 11 E0) and uses PowerPoint COM export + clean-deck generation when Python parsing/edit-in-place is not possible."
compatibility: "Windows. PowerPoint required for COM export/convert. Python optional for zip-backed PPTX workflows."
---

# pptx-enterprise

Enterprise-aware PowerPoint skill. Handles normal zip-backed .pptx files AND the Microsoft reality where many .pptx files on disk are actually OLE2 compound documents (legacy .ppt) or IRM / sensitivity-label-wrapped containers that Python cannot reliably parse or edit.

## A) Trigger rules

Use this skill any time:
- A `.pptx` file is referenced as input or output, OR
- The user mentions a deck, slides, presentation, pitch deck, slide deck, etc., OR
- The user asks to read, inspect, summarize, edit, add slides to, or build a PowerPoint.

## B) Safety / guardrails (non-negotiable)

- **Never remove or downgrade sensitivity / IRM labels** unless the user explicitly confirms it is permitted.
- **Never overwrite the source deck** without explicit user request AND only when the deck is zip-backed PPTX.
- **Default behavior for legacy/OLE2/IRM decks: do NOT edit in place.** Instead, generate a clean "drop-in" deck as output and instruct the user to copy/paste slides into the original if needed.
- Write all outputs under an `exports` folder next to the deck. If that location is not writable, fall back to `.\exports` in the current working directory.
- **No network calls.** This skill is fully local.
- Always release COM objects and terminate any PowerPoint process the skill itself started, so no orphaned `POWERPNT.EXE` is left behind.

## C) Detection workflow (MANDATORY FIRST STEP)

**Always detect before attempting Python.** In this tenant, Office files default to the "Confidential \ Internal Only" label and are frequently OLE2/IRM-wrapped on disk — `python-pptx` and `markitdown` will fail on these. Detect-first, then route:

- `zip` (header `50 4B ...` / "PK") → Python extraction is viable; if modules are missing, offer install OR fall back to COM export.
- `ole2` (header `D0 CF 11 E0`) → go straight to COM export. Do NOT attempt Python first.
- `unknown` → treat as OLE2; use COM.

Run:

```powershell
pwsh -File "%USERPROFILE%\.copilot\skills\pptx-enterprise\scripts\detect_container.ps1" -Path "<deck>"
```

Returns JSON: `{ path, headerHex, containerType }`.

- `containerType = "zip"` (header `50 4B ...`) → treat as a **normal zip-backed PPTX**. Python tooling and in-place edits are viable.
- `containerType = "ole2"` (header `D0 CF 11 E0`) → treat as **legacy .ppt OR IRM/sensitivity-wrapped OOXML**. Python parsing/editing is NOT reliable. Use PowerPoint COM.
- `containerType = "unknown"` → treat as legacy/wrapped; use COM.

## D) Read / Analyze workflows

### If container is `zip`:
1. Prefer Python extraction:
   ```powershell
   python -m markitdown "<deck>"
   ```
2. If `markitdown` is missing: either install it (`python -m pip install -U markitdown`) or fall back to the COM exporter.
3. Also run the COM exporter if the user wants slide PNGs (markitdown only returns text).

### If container is `ole2` / `unknown`:
1. Use the COM exporter (read-only, opens presentation with `ReadOnly=True`):
   ```powershell
   pwsh -File "%USERPROFILE%\.copilot\skills\pptx-enterprise\scripts\export_pptx_com.ps1" -Path "<deck>" -OutDir "<exports>"
   ```
   Produces `exports/slides.json` (text + notes) and `exports/slides/Slide-###.png`.

### After extraction (either path):
Run the analyzer:
```powershell
python "%USERPROFILE%\.copilot\skills\pptx-enterprise\scripts\analyze_exports.py" "<exportsDir>"
```

Produces:
- `exports/analysis_report.json`
- An exec-ready markdown summary printed to stdout.

## E) Editing workflows (change text, add slides)

### If the deck is zip-backed PPTX:
- You MAY edit in place **only if the user explicitly requests updating the existing file**.
- Preferred approach:
  1. If Python `python-pptx` is available and can open the file, use it for deterministic edits.
  2. If Python edits fail, fall back to PowerPoint COM, but warn the user about brittleness and confirm before writing.
- Otherwise, default to writing `exports/updated_deck.pptx` as a copy.

### If the deck is `ole2` / `unknown` (legacy / IRM-wrapped):
- **Default: generate a NEW clean deck at `exports/updated_deck.pptx`** containing the requested edits / new slides.
- Also export slide PNGs to `exports/slides/` for visual verification.
- Tell the user they can copy/paste the updated slides from the clean deck into the original deck (which preserves the sensitivity label).
- Optional conversion attempt (only with user permission, since it may strip the label):
  - Open in PowerPoint COM and `SaveCopyAs` to a new zip-backed PPTX at `exports/converted.pptx`.
  - Re-run `detect_container.ps1` on the output.
  - If the converted file has a `PK` header, edits-in-place can proceed on the converted copy (never on the original).

Editing / building uses:
```powershell
pwsh -File "%USERPROFILE%\.copilot\skills\pptx-enterprise\scripts\build_deck_com.ps1" `
  -OutPath "<exports>\updated_deck.pptx" `
  -SpecPath "<spec.json>" `
  [-BaseDeckPath "<deck.pptx>"]
```

`-BaseDeckPath` is used ONLY when that file is zip-backed; otherwise the script ignores it and creates a new deck.

## F) Build-from-scratch workflow

- Always create a new clean deck at `exports/new_deck.pptx` unless the user specifies another path.
- Follow **Executive Presentation Preferences** (if defined in your copilot-instructions):
  - Lean first draft; easier to add than cut.
  - **Slide 2 is always "The Ask"** (unless deck is purely educational).
  - Headline-style slide titles (stating the insight, not the topic).
  - Big stat callouts as the primary visual device.
  - Actions slide as a table: Action / Owner / Outcome.
  - Color palette: teal primary, terracotta `#C66B4E` accent, charcoal `#2E3A3F` text, ivory `#F5F1EA` background. No amber/orange.
  - Plain business language; no jargon.
- After building, export slide PNGs to `exports/slides/` for layout verification.

## G) Outputs contract

Every run should (where applicable) produce:

- `exports/slides.json` — per-slide text + speaker notes
- `exports/slides/Slide-###.png` — slide images
- `exports/analysis_report.json` — analyzer output
- For edit/build flows:
  - `exports/updated_deck.pptx` (edit flow), OR
  - `exports/new_deck.pptx` (build-from-scratch flow)

## Spec JSON format for build_deck_com.ps1

```json
{
  "slides": [
    {
      "layout": "title",
      "title": "Headline stating the insight",
      "subtitle": "Optional subtitle",
      "bullets": [],
      "notes": "Speaker notes for this slide."
    },
    {
      "layout": "content",
      "title": "Slide 2 — The Ask",
      "bullets": ["Approve X", "Fund Y", "Unblock Z"],
      "notes": ""
    }
  ]
}
```

Supported `layout` values: `title`, `content`, `blank`.

## H) Chart preflight gate + waterfall editing pattern

Office chart objects — especially **waterfall charts (ChartType=119)** — require a healthy Excel COM server to read or modify their embedded workbook. On many enterprise machines Excel COM can fail intermittently with `0x80080005 CO_E_SERVER_EXEC_FAILURE`, typically caused by stuck headless `EXCEL.EXE` processes or sign-in dialogs. Also critical: **waterfall chart *structure* is effectively read-only through COM** even when Excel COM is healthy. What works and what doesn't:

| Operation | Works via COM? |
|-----------|---------------|
| Read embedded workbook cells | ✅ Yes |
| Write embedded workbook cell values | ✅ Yes |
| `Slide.Duplicate()` (inherits chart formatting & subtotal flags) | ✅ Yes |
| `Presentation.Save()` (preserves IRM label) | ✅ Yes |
| `Chart.SetSourceData()` | ❌ HRESULT |
| `Series.XValues = ...` / `Values = ...` / `Formula = ...` | ❌ HRESULT |
| `Point(n).IsTotal = $true/$false` | ❌ HRESULT |

### Durable waterfall edit pattern
1. Run the preflight (below). If it fails, abort and surface remediation.
2. Duplicate the source slide (`Slide.Duplicate()`) — inherits chart formatting and subtotal flags at the SAME point indices as the source.
3. Activate the new slide's chart data (`Chart.ChartData.Activate()`), get the workbook, overwrite the cells with new category labels + values. **Cast `Cells.Item` row/col args to `[int]` and values to `[double]`** — bare PS integer arithmetic yields Doubles that blow up COM setters.
4. `Workbook.Close($true)` then `Presentation.Save()`.
5. **Tell the user up-front** that two things need a manual UI fix (~30 seconds): (a) range extension if the new data has a different row count than the source — right-click chart → Edit Data → drag the blue range boundary; (b) "Set as Total" toggles if the subtotal position changed — right-click bar → Set as Total / uncheck. Do NOT waste cycles trying to automate these; they don't work through COM.

### Preflight command

```powershell
pwsh -File "%USERPROFILE%\.copilot\skills\pptx-enterprise\scripts\Test-OfficeAutomationPreflight.ps1" `
  -DeckPath "<deck.pptx>" `
  [-RequireWaterfall] `
  [-Json]
```

The script:
1. **Sweeps headless `EXCEL.EXE` / `POWERPNT.EXE` orphans** (empty `MainWindowTitle`) before doing anything else.
2. Sniffs the deck header (`D0 CF 11 E0` = ole2/IRM, `50 4B` = zip).
3. Instantiates `PowerPoint.Application` via COM and opens the deck read-only.
4. Does NOT explicitly instantiate Excel. With `-RequireWaterfall`, it locates the first `Chart.ChartType=119` chart and calls `Chart.ChartData.Activate()` — this lets PowerPoint spawn Excel internally, which is the reliable path on this machine. Verifies a real `Workbook` object is returned.
5. Always tears down presentations, workbooks, and applications it created; also kills any Office PIDs that did not exist before the run.

### Exit codes

| Code | Meaning |
|------|---------|
| 0 | All required checks passed. |
| 2 | Bad arguments or file-not-found. |
| 3 | PowerPoint COM could not instantiate or open the deck. |
| 5 | Waterfall required but `ChartData.Activate()` failed (the 0x80080005 environmental failure surfaces here). |
| 6 | Waterfall required but no waterfall chart found in the deck. |

### When callers must invoke this

- Any read/edit flow targeting a deck that is OLE2/IRM-wrapped **and** contains charts.
- Any edit flow that adds, modifies, or reads data in a waterfall chart. Pass `-RequireWaterfall`.
- Any build flow that inserts waterfall charts (even into a new deck). Pass `-RequireWaterfall` against a representative reference deck.

### On preflight failure

If the preflight returns a non-zero exit code, the skill MUST:

1. Stop immediately. Do not open, modify, or save the deck.
2. Emit the remediation checklist verbatim from the script output.
3. Refuse to draw shapes, export images, or otherwise approximate a chart.
4. Wait for the user to remediate and re-run.

### COM shape mutation rule

For rotated shapes (e.g., `Left Brace 18`, `Right Brace`, any shape with `Rotation != 0`), do NOT resize or reposition by mutating `Left/Top/Width/Height/Rotation` properties — the rotation math does not compose cleanly via COM and the shape will distort. Instead, `Copy()` the known-good shape from a source slide and `Paste()` onto the target slide; reposition the pasted copy as a single unit only if necessary.

### Save rule for IRM files

For decks whose container type is `ole2` (IRM/sensitivity-labeled), use `Presentation.Save()` to write in place. Do NOT use `SaveAs` with `ppSaveAsOpenXMLPresentation=24` — it can re-wrap or mutate the IRM container and may strip or downgrade the sensitivity label.
