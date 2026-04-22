---
name: docx-enterprise
purpose: IRM-aware Word wrapper that mirrors pptx-enterprise / excel-enterprise
---

# docx-enterprise (README)

IRM / sensitivity-label-aware Word skill for enterprise tenants where `.docx` files default to
"Confidential \ Internal Only" and are frequently OLE2-wrapped on disk. `python-docx` cannot
parse these; this skill routes to Word COM via a detect-then-route pattern.

## Layout

- `SKILL.md` — the operational skill spec (trigger rules, workflows, capability matrix).
- `scripts/detect_container.ps1` — magic-byte sniff (`50 4B` zip vs `D0 CF 11 E0` OLE2).
- `scripts/export_docx_com.ps1` — read-only Word COM exporter. Produces `document.json` (paragraphs, revisions, comments, structure) and `document.txt`.
- `scripts/build_doc_com.ps1` — Word COM write driver. Supported modes: `accept-changes`, `find-replace`, `add-comment`, `extract-redlines`.
- `scripts/analyze_export.py` — post-extraction summarizer. Produces a short markdown report (headings, revision stats by author, open comments).

## Design principles

1. **Detect before acting.** Never call `python-docx` on a file without first confirming it is zip-backed.
2. **Default to `Document.Save()`** on IRM docs. `SaveAs2` can mutate the IRM container and strip the label.
3. **COM hygiene matters.** Always Quit + force-kill spawned WINWORD.EXE; never kill a user-visible window.
4. **Read-only by default.** Edit modes are opt-in via explicit flags in `build_doc_com.ps1`.

## Companion skills

- `excel-enterprise` — same pattern for `.xlsx`.
- `pptx-enterprise` — same pattern for `.pptx`.

The three skills share the same header sniff convention so routing decisions are consistent across Office formats.
