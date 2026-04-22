# CLI Buddy Starter

One-command setup for **GitHub Copilot CLI** on a Windows machine, tuned for
people who work with sensitivity-labeled (IRM) Microsoft Office files.

Runs in ~2 minutes. No admin, no registry changes, no background services.

## What you get

Five things in one installer:

| # | Category | What | How it's installed |
|---|---|---|---|
| 1 | **Enterprise skills** | IRM-aware PowerPoint / Word / Excel handling (`pptx-enterprise`, `docx-enterprise`, `excel-enterprise`) | `install.ps1` drops folders into `~/.copilot/skills/` |
| 2 | **Anthropic skills** | Generic `pptx`, `docx`, `pdf`, `xlsx` from [anthropics/skills](https://github.com/anthropics/skills) | `gh skill install` |
| 2b | **Community skills** | `excel-toolkit`, `writing-plans` (from [Sentry01](https://github.com/Sentry01/copilot-cli-skills)); `meeting-prep`, `project-status`, `research` (from [jimbanach @ v1.5.1](https://github.com/jimbanach/copilot-cli-starter)) | `gh skill install` |
| 3 | **Copilot plugins** | `microsoft-docs`, `power-bi-development`, `workiq` | `copilot plugin install` |
| 4 | **Instructions snippet** | Optional guidance block that teaches Copilot to route Office files correctly | Manual copy-paste |

Categories 1-3 are done by the installer. Category 4 is a 30-second copy-paste.

## Mental model (keep these straight)

```
install.ps1         -> your enterprise skills      (no auth)
gh skill install    -> community / Anthropic skills (needs gh auth)
copilot plugin      -> live integrations           (needs copilot CLI)
snippet             -> optional advice             (manual)
```

If someone uses the wrong command for the wrong thing, they'll get cryptic
errors. The installer keeps them straight.

## Prerequisites

- Windows 10 / 11
- **GitHub CLI** (`gh`) installed - https://cli.github.com/
- **GitHub Copilot CLI** installed - https://docs.github.com/en/copilot/concepts/agents/about-copilot-cli
- **GitHub CLI authenticated** - if you haven't already:
  ```powershell
  gh auth login --hostname github.com
  ```
  Choose **HTTPS** and **Login with a web browser**. One-time setup.
- Microsoft Office installed (Word / Excel / PowerPoint) - the enterprise
  skills' COM fallback needs it
- Python 3.9+ (optional, used for the fast path on non-IRM files)

## Install (one line)

Open PowerShell and run:

```powershell
iwr https://raw.githubusercontent.com/nfadorsen/cli-buddy-starter/main/install.ps1 | iex
```

The installer prints what it's doing section by section. Each section either
succeeds, fails with a clear message, or is skipped (e.g., if `gh` isn't
authenticated yet). Re-running is safe.

At the end you'll see a summary like:

```
== Summary
Enterprise skills      : 3 ok, 0 failed
Anthropic skills       : 4 ok, 0 failed
Copilot plugins        : 3 ok, 0 failed
Instructions snippet   : manual (see next steps)
```

## Step 4 (manual) - add the instructions snippet

This is optional but recommended. It makes Copilot CLI behave more
predictably around Office files.

1. Open this file in your browser:
   https://github.com/nfadorsen/cli-buddy-starter/blob/main/copilot-instructions.snippet.md
2. Click **Raw** (top-right of the file view).
3. Copy **everything** - including the `<!-- BEGIN ... -->` and
   `<!-- END ... -->` comment lines.
4. Open your repo's `.github/copilot-instructions.md` (create it if missing).
5. Paste **at the end** of the file. Don't replace anything above.
6. Save. Next Copilot CLI session picks it up automatically.

### To update the snippet later

Find the block between the `BEGIN` and `END` markers in your instructions
file, delete it, and paste the new version. Nothing else in your file is
touched.

### To remove

Delete everything between the `BEGIN` and `END` markers (including the
markers). Your original file is back to exactly what it was.

## Verify

Start a Copilot CLI session and run:

```
/skills
```

You should see the enterprise and Anthropic skills listed. Then:

```
/env
```

Shows loaded instructions, MCP servers, skills, agents, and plugins - use
this to confirm everything is wired up.

## Customize / skip sections

The installer takes optional parameters. Run locally (not via `iwr | iex`)
to pass them:

```powershell
# Download once and run with custom args
iwr https://raw.githubusercontent.com/nfadorsen/cli-buddy-starter/main/install.ps1 `
  -OutFile $env:TEMP\install.ps1
& $env:TEMP\install.ps1 -Skip anthropic,plugins -Force
```

Parameters:

| Parameter | Default | Purpose |
|---|---|---|
| `-Skip` | `none` | `enterprise`, `anthropic`, `community`, `plugins`, `snippet`, `all`, or `none` (combine with commas) |
| `-Force` | off | Overwrite existing enterprise skill folders |
| `-EnterpriseSkills` | `pptx-,docx-,excel-enterprise` | Which enterprise skills to install |
| `-AnthropicSkills` | `pptx, docx, pdf, xlsx` | Which Anthropic skills to install |
| `-SentrySkills` | `excel-toolkit, writing-plans` | Skills from Sentry01/copilot-cli-skills |
| `-JimbanachSkills` | `meeting-prep, project-status, research` | Skills from jimbanach/copilot-cli-starter |
| `-JimbanachRef` | `v1.5.1` | Pin for the jimbanach skill repo |
| `-Plugins` | `microsoft-docs@awesome-copilot, power-bi-development@awesome-copilot, workiq@copilot-plugins` | Which plugins to install |

## Safety properties

- No admin / elevation required
- No telemetry. Network only to `github.com` (install) and local Office COM (runtime)
- Enterprise skills never remove or downgrade sensitivity labels
- Enterprise skills open source files `ReadOnly` by default; edits are explicit
- Output is written to an `exports\` folder next to the source file
- Every section is opt-out via `-Skip`

## Uninstall

```powershell
# Enterprise skills
Remove-Item "$env:USERPROFILE\.copilot\skills\pptx-enterprise"  -Recurse -Force
Remove-Item "$env:USERPROFILE\.copilot\skills\docx-enterprise"  -Recurse -Force
Remove-Item "$env:USERPROFILE\.copilot\skills\excel-enterprise" -Recurse -Force

# Anthropic skills (from inside Copilot CLI)
#   /skills remove <path shown by /skills list>

# Plugins (from inside Copilot CLI or CLI)
copilot plugin uninstall microsoft-docs@awesome-copilot
copilot plugin uninstall power-bi-development@awesome-copilot
copilot plugin uninstall workiq@copilot-plugins

# Instructions snippet
#   delete everything between BEGIN / END markers in .github/copilot-instructions.md
```

## Troubleshooting

**`gh skill install` says I'm not authenticated.**
Run `gh auth login --hostname github.com`, pick HTTPS + web browser, then
re-run the installer.

**A plugin install fails with "plugin not found".**
The marketplace name matters. `microsoft-docs@copilot-plugins` doesn't
exist - it's `microsoft-docs@awesome-copilot`. Two marketplaces ship by
default:
- `copilot-plugins` hosts `workiq`, `spark`, `advanced-security`
- `awesome-copilot` hosts most community plugins including `microsoft-docs`
  and `power-bi-development`

**`/skills add <url>` doesn't work.**
Correct - `/skills add` takes a **local directory**, not a URL. Use
`gh skill install <repo> <name> --scope user` instead. The installer does
this for you.

**My enterprise skill isn't loading after re-install.**
Restart your Copilot CLI session. Skills load at session start.

## Support

File an issue on this repo, or ping the maintainer.
