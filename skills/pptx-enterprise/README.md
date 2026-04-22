# pptx-enterprise

User-scope Copilot skill for working with PowerPoint decks in an enterprise environment where IRM / sensitivity labels and legacy OLE2 files are common.

## Why this skill exists

A `.pptx` file on disk is not always a modern zip-backed Office Open XML file. In this environment you will routinely encounter:

- **Zip-backed PPTX** — first bytes `50 4B 03 04` ("PK"). Modern, parseable by `python-pptx`, `markitdown`, etc.
- **OLE2 compound document** — first bytes `D0 CF 11 E0`. This can be either:
  - Legacy binary `.ppt` saved with a `.pptx` extension, OR
  - Modern PPTX wrapped in an IRM / sensitivity-label container (encrypted OOXML inside an OLE2 shell).

Python PPTX tooling **cannot reliably parse OLE2-wrapped files**, and attempts to save them as a clean PPTX typically re-wrap back into OLE2 because the sensitivity label is preserved by Office.

## The safe enterprise editing model

1. **Detect first.** `detect_container.ps1` reads the magic bytes and classifies the file as `zip`, `ole2`, or `unknown`.
2. **Read via COM for OLE2.** `export_pptx_com.ps1` uses PowerPoint itself to export slide text, speaker notes, and PNG images.
3. **Edit clean, never wrapped.** For OLE2/IRM decks, default to producing a separate clean "drop-in" deck (`exports/updated_deck.pptx`). The user copies/pastes slides into the original, which preserves the sensitivity label.
4. **Never strip labels** unless the user explicitly confirms it is permitted.

## How to test

Run these commands in PowerShell after the skill is installed:

```powershell
# 1. Detect
pwsh -File "$env:USERPROFILE\.copilot\skills\pptx-enterprise\scripts\detect_container.ps1" -Path "C:\path\to\deck.pptx"

# 2. Export (COM, read-only)
pwsh -File "$env:USERPROFILE\.copilot\skills\pptx-enterprise\scripts\export_pptx_com.ps1" -Path "C:\path\to\deck.pptx" -OutDir "C:\path\to\exports"

# 3. Analyze
python "$env:USERPROFILE\.copilot\skills\pptx-enterprise\scripts\analyze_exports.py" "C:\path\to\exports"

# 4. Build from spec
pwsh -File "$env:USERPROFILE\.copilot\skills\pptx-enterprise\scripts\build_deck_com.ps1" -OutPath "C:\path\to\exports\new_deck.pptx" -SpecPath "C:\path\to\spec.json"
```

Expected outputs in the `exports` folder:

- `slides.json`
- `slides/Slide-001.png`, `slides/Slide-002.png`, ...
- `analysis_report.json`
- `updated_deck.pptx` or `new_deck.pptx` when editing/building

## Requirements

- Windows + installed Microsoft PowerPoint (COM)
- PowerShell 5.1 or pwsh 7+
- Python 3.x for the analyzer (standard library only; no network deps)
- Optional: `pip install markitdown python-pptx` for zip-backed PPTX workflows
