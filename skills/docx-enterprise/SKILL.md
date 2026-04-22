---
name: docx-enterprise
description: "Enterprise Word workflow. Reads, analyzes, and edits .docx files. Detects IRM/sensitivity-label-wrapped OLE2 containers (D0 CF 11 E0) and uses Word COM when python-docx cannot parse the file. Preserves tenant sensitivity labels on save."
compatibility: "Windows. Microsoft Word required for COM path. Python + python-docx optional for zip-backed .docx workflows."
---

# docx-enterprise

Enterprise-aware Word skill. Handles normal zip-backed `.docx` files AND the Microsoft reality where many `.docx` files on disk (contracts, memos, policy documents) are OLE2-wrapped IRM / sensitivity-labeled containers that `python-docx` cannot parse (throws `BadZipFile`).

In many enterprise tenants, Office files default to a sensitivity label (e.g., "Confidential") that routinely produces OLE2-wrapped `.docx` on disk. Contracts, HR/legal memos, and policy documents are the high-exposure file class for this path.

## A) Trigger rules

Use this skill any time:
- A `.docx` file is referenced as input or output, OR
- The user mentions a Word document, contract, MSA, memo, letter, or policy document, OR
- The user asks to read, review, accept tracked changes, add comments, or find-and-replace in a Word document.

Do **not** use this skill for plain text files, Markdown, or PDFs.

## B) Safety / guardrails (non-negotiable)

- **Never remove or downgrade sensitivity / IRM labels** unless the user explicitly confirms it is permitted.
- **Never overwrite the source document** without explicit user request AND only when the document is zip-backed `.docx`.
- **Default behavior for OLE2/IRM documents: do NOT edit in place.** Read via COM read-only; write changes to a new file or to a clean copy the user can paste into the original.
- Save-rule for OLE2/IRM documents: use `Document.Save()` to write in place (preserves the label). Do NOT `SaveAs` to a new format — it can re-wrap or mutate the IRM container and strip the label.
- Write all outputs under an `exports` folder next to the source document. If that location is not writable, fall back to `.\exports` in the current working directory.
- **No network calls.** This skill is fully local.
- Always release COM objects and terminate any Word process the skill itself started, so no orphaned `WINWORD.EXE` is left behind.

## C) Detection workflow (MANDATORY FIRST STEP)

**Always detect before attempting Python.** `python-docx` opens `.docx` via `ZipFile`; an IRM-wrapped docx has an OLE2 header and will throw `BadZipFile`. Detect-first avoids the wasted round-trip.

Run:

```powershell
pwsh -File "%USERPROFILE%\.copilot\skills\docx-enterprise\scripts\detect_container.ps1" -Path "<doc.docx>"
```

Returns JSON: `{ path, headerHex, containerType }`.

- `containerType = "zip"` (header `50 4B ...`) → normal zip-backed `.docx`. `python-docx` / `markitdown` are viable.
- `containerType = "ole2"` (header `D0 CF 11 E0`) → IRM/sensitivity-wrapped OOXML. Python parsing will fail. Use Word COM.
- `containerType = "unknown"` → treat as OLE2; use COM.

## D) Read / Analyze workflows

### If container is `zip`:
1. Prefer Python extraction:
   ```powershell
   python -m markitdown "<doc.docx>"
   ```
2. Or `python-docx` for programmatic access.
3. Use the COM exporter below if the user also wants tracked-changes / comments structure preserved in JSON.

### If container is `ole2` / `unknown`:
Use the COM exporter (read-only, `ReadOnly=True`):

```powershell
pwsh -File "%USERPROFILE%\.copilot\skills\docx-enterprise\scripts\export_docx_com.ps1" -Path "<doc.docx>" -OutDir "<exports>"
```

Produces:
- `exports/document.json` — paragraphs (with style), headings, revisions (tracked changes), comments, section + page count
- `exports/document.txt` — plain text of the body for quick reading / grep

### After extraction (either path):
Run the analyzer:

```powershell
python "%USERPROFILE%\.copilot\skills\docx-enterprise\scripts\analyze_export.py" "<exportsDir>"
```

Produces a short markdown summary (headings, revision count by author, open comment threads, notable risk language) printed to stdout.

## E) Editing workflows

### If the document is zip-backed `.docx`:
- You MAY edit in place **only if the user explicitly requests updating the existing file**.
- Use `python-docx` for deterministic edits (find/replace, paragraph insertion, style changes).
- If Python edits fail or the document has complex content (tracked changes, SDTs, nested tables), fall back to Word COM.

### If the document is `ole2` / `unknown` (IRM-wrapped) — USE COM:
Use the build/edit driver:

```powershell
pwsh -File "%USERPROFILE%\.copilot\skills\docx-enterprise\scripts\build_doc_com.ps1" `
  -Path "<doc.docx>" `
  -Mode <accept-changes|find-replace|add-comment|extract-redlines> `
  [-Find "<pattern>" -Replace "<replacement>"] `
  [-Anchor "<anchor text>" -CommentText "<comment>" -Author "<initials>"] `
  [-OutPath "<optional new path>"]
```

Supported modes:

| Mode               | Purpose                                                             | Save behavior                               |
|--------------------|---------------------------------------------------------------------|---------------------------------------------|
| `accept-changes`   | Accept all tracked changes in the document                          | `Save()` in place (preserves label)         |
| `find-replace`     | Exact-match find/replace across body, headers, footers              | `Save()` in place (preserves label)         |
| `add-comment`      | Add a comment anchored to the first match of `-Anchor`              | `Save()` in place (preserves label)         |
| `extract-redlines` | Export every revision as structured JSON (no write)                 | Read-only; writes `exports/redlines.json`   |

**If `-OutPath` is provided**, the driver copies the file first and edits the copy (safer default for first use).

### Capability matrix (what works / what doesn't via COM)

| Operation                                  | Reliable via Word COM? |
|--------------------------------------------|------------------------|
| Read body text                             | ✅ Yes                 |
| Read `Revisions` collection                | ✅ Yes                 |
| Read `Comments` collection                 | ✅ Yes                 |
| `Revisions.AcceptAll()`                    | ✅ Yes                 |
| `Selection.Find.Execute` with Replace      | ✅ Yes                 |
| `Comments.Add(Range, Text)`                | ✅ Yes                 |
| `Document.Save()` (preserves IRM label)    | ✅ Yes                 |
| `Document.SaveAs2()` on IRM doc            | ⚠️ Risks label mutation — avoid unless user confirms |

## F) COM hygiene

- Word supports full `Application.Visible = $false` (unlike PowerPoint). Use it.
- Wrap every COM session in try/finally; always call `Documents.Close($false)` then `Application.Quit()`, then force-kill any `WINWORD.EXE` the skill spawned.
- Track the Word PIDs that existed before starting; on cleanup, kill only new PIDs. Never kill a user-visible Word window.
- Marshaling: when mutating ranges or passing numeric args to Word COM, cast to the expected types (`[int]`, `[string]`). PowerShell's default numeric is Double, which Word COM setters frequently reject.

## G) Error handling

- If Word COM fails to instantiate (`0x80080005` or similar), abort — do NOT fall back to pretending the document was read. Surface the error to the user with remediation: close Word, end orphaned `WINWORD.EXE` processes, open Word once interactively, then retry.
- If the sensitivity label cannot be read (Word < 2019 in tenant), proceed but warn the user that label preservation on save is not independently verified; advise them to confirm via File → Info after the save.
- If `detect_container.ps1` reports `unknown`, default to the COM path and log the header bytes so the user can investigate.
