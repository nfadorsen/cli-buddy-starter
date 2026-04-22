# CLI Buddy Starter

Enterprise-aware Copilot CLI skills for Windows users whose Office files are
protected by **Microsoft Information Protection (IRM / sensitivity labels)**.

If you've ever tried to have Copilot read a `.pptx`, `.docx`, or `.xlsx` only
to have it fail with `BadZipFile` or silently return nothing — this is the fix.

## What's included

Three local skills that plug into the GitHub Copilot CLI:

| Skill | What it does |
|---|---|
| `pptx-enterprise` | Reads, analyzes, edits, and creates PowerPoint decks. Detects IRM/OLE2-wrapped containers and uses PowerPoint COM when Python can't parse the file. Preserves sensitivity labels on save. |
| `docx-enterprise` | Reads, analyzes, and edits Word docs. Detects IRM/OLE2-wrapped containers and uses Word COM when `python-docx` fails. Preserves sensitivity labels. |
| `excel-enterprise` | Reads and analyzes workbooks. Detects IRM/OLE2-wrapped workbooks and falls back to Excel COM export (CSV) for analysis. Never modifies the source workbook. |

All three skills follow a **detect-first** pattern: they sniff the file's
magic bytes and only use COM when Python tooling genuinely can't read the file.

## Prerequisites

- Windows 10 / 11
- GitHub Copilot CLI already installed (`gh copilot` available in a terminal)
- Microsoft Office installed (Word / Excel / PowerPoint) — required for the
  COM fallback path on IRM-wrapped files
- Python 3.9+ (optional, used for the fast path on non-IRM files)

## Install

Setup has two steps. Both are copy-paste, both use native mechanisms, both are reversible.

### Step 1 — Install the enterprise skills (PowerShell, no admin)

```powershell
iwr https://raw.githubusercontent.com/nfadorsen/cli-buddy-starter/main/install.ps1 | iex
```

This downloads `pptx-enterprise`, `docx-enterprise`, and `excel-enterprise` into
`%USERPROFILE%\.copilot\skills\`. No admin, no registry changes, no scheduled
tasks, no telemetry. [See the installer script for full details](./install.ps1).

### Step 2 — (Optional) Add extra GitHub-hosted skills

If you also want `pptx`, `docx`, `pdf`, `meeting-prep`, `project-status`, and
`research` (general-purpose skills, not IRM-aware), open a Copilot CLI session
and run:

```
/skills
```

Add this as a skill source:

```
https://github.com/jimbanach/copilot-cli-starter
```

(pin to tag `v1.5.1` if prompted). Copilot CLI will manage installs and
updates from there natively — you don't need to re-run anything from this repo
to keep them current.

> Copilot CLI ships with `writing-plans` and `excel-toolkit` built in, so you
> don't need to install those separately.

### Verify

Inside a Copilot CLI session, run `/skills` (or `/env`) to confirm all the
expected skills are loaded.

## Uninstall

Delete the skill folders:

```powershell
Remove-Item "$env:USERPROFILE\.copilot\skills\pptx-enterprise"  -Recurse -Force
Remove-Item "$env:USERPROFILE\.copilot\skills\docx-enterprise"  -Recurse -Force
Remove-Item "$env:USERPROFILE\.copilot\skills\excel-enterprise" -Recurse -Force
```

## Try it

Open a Copilot CLI session in any folder and ask:

> "Inspect `<some-file>.pptx` and describe the structure."

The CLI will auto-pick `pptx-enterprise` when the file is a `.pptx`.

## Optional: custom instructions

`copilot-instructions.sample.md` contains a sanitized set of enterprise-
software defaults (Office COM hygiene, IRM detection, safety rails, tone).
Copy it into your own repo at `.github/copilot-instructions.md` and edit to
taste. It's independent of the skills — the skills work without it.

## Safety properties

- No admin / elevation required
- No telemetry, no network beyond GitHub (install) and local Office (runtime)
- Scripts never remove or downgrade sensitivity labels
- Scripts open source files `ReadOnly` by default; edits are explicit and opt-in
- Output is written to an `exports\` folder next to the source file

## Support

File an issue on this repo, or ping the maintainer.
