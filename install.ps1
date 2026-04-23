# CLI Buddy Starter - one-shot setup for Copilot CLI
#
# Four things, idempotent, fail-tolerant per section:
#   1. Enterprise skills       -> ~/.copilot/skills/ (no auth)
#   2. Anthropic doc skills    -> gh skill install (requires gh auth)
#   3. Copilot CLI plugins     -> copilot plugin install
#   4. Instructions snippet    -> printed next steps (manual paste)
#
# Safety properties:
#   - No admin / elevation
#   - No registry changes, no scheduled tasks, no background services
#   - No telemetry. Only reaches github.com (clone/API) and your Copilot CLI.
#   - Every section prints what it's doing and whether it succeeded.
#   - Re-runnable. Use -Force to overwrite enterprise skills.
#
# Skip sections you don't want:
#   iwr .../install.ps1 | iex
#   .\install.ps1 -Skip anthropic,plugins     # just the enterprise skills

[CmdletBinding()]
param(
    [string]$Repo   = 'nfadorsen/cli-buddy-starter',
    [string]$Branch = 'main',
    [string[]]$EnterpriseSkills = @('pptx-enterprise','docx-enterprise','excel-enterprise'),
    [string[]]$AnthropicSkills  = @('pptx','docx','pdf','xlsx'),
    [string[]]$SentrySkills     = @('excel-toolkit','writing-plans'),
    [string[]]$JimbanachSkills  = @('research'),
    [string]$JimbanachRef       = 'v1.5.1',
    [string[]]$Plugins = @(
        'microsoft-docs@awesome-copilot',
        'power-bi-development@awesome-copilot',
        'workiq@copilot-plugins'
    ),
    [ValidateSet('enterprise','anthropic','community','plugins','snippet','all','none')]
    [string[]]$Skip = @('none'),
    [switch]$Force,
    [switch]$AddSnippet,
    [switch]$SetExecutionPolicy,
    [switch]$InstallMissing
)

$ErrorActionPreference = 'Stop'

# ------------- helpers -------------
function Info($m) { Write-Host "  $m" -ForegroundColor Cyan }
function Ok($m)   { Write-Host "  [ok]  $m" -ForegroundColor Green }
function Warn2($m){ Write-Host "  [!!]  $m" -ForegroundColor Yellow }
function Fail2($m){ Write-Host "  [xx]  $m" -ForegroundColor Red }
function Section($n,$t) {
    Write-Host ""
    Write-Host "== Step $n :: $t" -ForegroundColor White
}
function HasCommand($name) {
    return [bool](Get-Command $name -ErrorAction SilentlyContinue)
}
function InSkip($name) { return ($Skip -contains $name) -or ($Skip -contains 'all') }

$summary = [ordered]@{
    'Enterprise skills' = 'skipped'
    'Anthropic skills'  = 'skipped'
    'Community skills'  = 'skipped'
    'Copilot plugins'   = 'skipped'
    'Instructions snippet' = 'manual (see next steps)'
}
$failures = New-Object System.Collections.Generic.List[object]
function AddFailure($section, $item, $reason, $retry) {
    $failures.Add([pscustomobject]@{
        Section = $section; Item = $item; Reason = "$reason".Trim(); Retry = $retry
    })
}

# One-line descriptions shown during install so colleagues see what each thing does
$descriptions = @{
    'pptx-enterprise'                      = 'IRM-safe PowerPoint automation via COM'
    'docx-enterprise'                      = 'IRM-safe Word automation via COM'
    'excel-enterprise'                     = 'IRM-safe Excel automation via COM'
    'pptx'                                 = 'generic PowerPoint creation/editing'
    'docx'                                 = 'generic Word document handling'
    'pdf'                                  = 'PDF read, merge, split, OCR'
    'xlsx'                                 = 'generic Excel spreadsheet handling'
    'excel-toolkit'                        = 'Excel read/edit/analyze helpers'
    'writing-plans'                        = 'plan and spec scaffolding before coding'
    'research'                             = 'structured research (MS Learn + web + WorkIQ)'
    'meeting-prep'                         = 'meeting agenda and talking-points workflow'
    'project-status'                       = 'project status report generation'
    'microsoft-docs@awesome-copilot'       = 'live Microsoft Learn documentation search'
    'power-bi-development@awesome-copilot' = 'Power BI / DAX guidance and review'
    'workiq@copilot-plugins'               = 'M365 workplace intelligence (email, meetings, Teams)'
}
function Desc($name) {
    if ($descriptions.ContainsKey($name)) { return " - $($descriptions[$name])" }
    return ''
}

Write-Host ""
Write-Host "CLI Buddy Starter - Copilot CLI setup" -ForegroundColor White
Write-Host "Source: https://github.com/$Repo (branch: $Branch)" -ForegroundColor DarkGray

# ------------- pre-flight -------------
Section 0 "Pre-flight checks"

function Refresh-PathFromEnv {
    # After winget installs a CLI, the new PATH isn't visible to the running process.
    # Re-read Machine + User scope PATH from the registry so subsequent HasCommand calls see the tool.
    $machinePath = [System.Environment]::GetEnvironmentVariable('Path','Machine')
    $userPath    = [System.Environment]::GetEnvironmentVariable('Path','User')
    $env:Path = "$machinePath;$userPath"
}

function Try-WingetInstall {
    param([string]$PackageId, [string]$FriendlyName)
    if (-not (HasCommand winget)) {
        Warn2 "winget not available - cannot auto-install $FriendlyName."
        Warn2 "Install manually: see https://learn.microsoft.com/windows/package-manager/winget/"
        return $false
    }
    Info "Installing $FriendlyName via winget (this can take ~30-60 sec)"
    try {
        & winget install --id $PackageId -e --accept-package-agreements --accept-source-agreements --silent | Out-Host
        if ($LASTEXITCODE -eq 0) {
            Ok "$FriendlyName installed"
            Refresh-PathFromEnv
            return $true
        }
        Fail2 "$FriendlyName winget install exited $LASTEXITCODE"
        return $false
    } catch {
        Fail2 "$FriendlyName winget install error: $($_.Exception.Message)"
        return $false
    }
}

$hasGit      = HasCommand git
$hasGh       = HasCommand gh
$hasCopilot  = HasCommand copilot

# Auto-install git and gh if missing (both needed by Step 3 plugins and Step 2/2b skills respectively).
$missingTools = @()
if (-not $hasGit) { $missingTools += @{ Id = 'Git.Git';    Name = 'git';    Cmd = 'git' } }
if (-not $hasGh)  { $missingTools += @{ Id = 'GitHub.cli'; Name = 'gh CLI'; Cmd = 'gh'  } }

if ($missingTools.Count -gt 0) {
    foreach ($t in $missingTools) { Warn2 "$($t.Name) not found" }
    $doInstall = $false
    if ($InstallMissing) {
        $doInstall = $true
        Info "Installing missing tools (per -InstallMissing)"
    } else {
        try {
            $names = ($missingTools | ForEach-Object { $_.Name }) -join ', '
            $answer = Read-Host "Install missing tools ($names) via winget now? (no admin) [y/N]"
            $doInstall = ($answer -match '^(y|yes)$')
        } catch {
            Warn2 "Non-interactive session - skipping auto-install. Re-run with -InstallMissing."
        }
    }
    if ($doInstall) {
        foreach ($t in $missingTools) {
            if (-not (HasCommand $t.Cmd)) {
                Try-WingetInstall -PackageId $t.Id -FriendlyName $t.Name | Out-Null
            }
        }
        # Re-detect after installs
        $hasGit = HasCommand git
        $hasGh  = HasCommand gh
    }
}

$ghAuthed = $false
if ($hasGh) {
    try {
        gh auth status 2>&1 | Out-Null
        $ghAuthed = ($LASTEXITCODE -eq 0)
    } catch { $ghAuthed = $false }
}

if ($hasGit)     { Ok "git found" }               else { Warn2 "git not found (needed for plugin installs)" }
if ($hasGh)      { Ok "gh CLI found" }            else { Warn2 "gh CLI not found (needed for Anthropic skills)" }
if ($ghAuthed)   { Ok "gh CLI authenticated" }    elseif ($hasGh) { Warn2 "gh CLI not authenticated (run: gh auth login --hostname github.com)" }
if ($hasCopilot) { Ok "copilot CLI found" }       else { Warn2 "copilot CLI not found (needed for plugins)" }

# Execution policy check - Windows default 'Restricted' blocks copilot.ps1 (used by plugin installs).
# Fix: Set CurrentUser scope to RemoteSigned. No admin required, safe default for dev machines.
$epEffective   = Get-ExecutionPolicy
$epBlocking    = @('Restricted','AllSigned','Undefined')
$epIsBlocking  = $epEffective -in $epBlocking
$epGpoForced   = $false
if ($epIsBlocking) {
    try {
        $epMachine = Get-ExecutionPolicy -Scope MachinePolicy
        $epUser    = Get-ExecutionPolicy -Scope UserPolicy
        if (($epMachine -in @('Restricted','AllSigned')) -or ($epUser -in @('Restricted','AllSigned'))) {
            $epGpoForced = $true
        }
    } catch { $epGpoForced = $false }
}

if ($epIsBlocking) {
    Warn2 "PowerShell execution policy is '$epEffective' - this blocks 'copilot' plugin installs"
    if ($epGpoForced) {
        Warn2 "Policy is enforced by Group Policy - cannot override without IT. Plugins (Step 3) will be skipped."
    } else {
        $doSet = $false
        if ($SetExecutionPolicy) {
            $doSet = $true
            Info "Setting execution policy (per -SetExecutionPolicy)"
        } else {
            try {
                $answer = Read-Host "Set CurrentUser execution policy to 'RemoteSigned' now? (no admin, recommended) [y/N]"
                $doSet  = ($answer -match '^(y|yes)$')
            } catch {
                Warn2 "Non-interactive session - skipping auto-fix. Plugins (Step 3) will fail."
                Warn2 "Re-run with -SetExecutionPolicy, or run manually: Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned"
            }
        }
        if ($doSet) {
            try {
                Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force -ErrorAction Stop
                Ok "Set CurrentUser execution policy to RemoteSigned"
                $epIsBlocking = $false
            } catch {
                Fail2 "Could not set execution policy: $($_.Exception.Message)"
                Warn2 "Run manually: Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned"
            }
        } else {
            Warn2 "Skipping execution policy change. Plugins (Step 3) will fail with 'running scripts is disabled'."
            Warn2 "To fix later: Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned"
        }
    }
} else {
    Ok "PowerShell execution policy OK ($epEffective)"
}

# ------------- 1. Enterprise skills -------------
Section 1 "Enterprise skills (~/.copilot/skills)"

if (InSkip 'enterprise') {
    Warn2 "Skipping enterprise skills (per -Skip)"
} else {
    $skillsRoot = Join-Path $env:USERPROFILE '.copilot\skills'
    if (-not (Test-Path $skillsRoot)) {
        New-Item -ItemType Directory -Path $skillsRoot -Force | Out-Null
        Info "Created $skillsRoot"
    }

    function Download-SkillFolder {
        param([string]$Skill)
        $apiUrl = "https://api.github.com/repos/$Repo/contents/skills/$Skill" + "?ref=$Branch"
        $target = Join-Path $skillsRoot $Skill
        Info "Installing $Skill$(Desc $Skill)"
        if ((Test-Path $target) -and -not $Force) {
            Warn2 "$Skill already exists (use -Force to overwrite). Skipping."
            return $true
        }
        if (Test-Path $target) { Remove-Item $target -Recurse -Force }
        New-Item -ItemType Directory -Path $target -Force | Out-Null
        $stack = New-Object System.Collections.Stack
        $stack.Push(@{ Url = $apiUrl; Rel = '' })
        while ($stack.Count -gt 0) {
            $item = $stack.Pop()
            $resp = Invoke-RestMethod -Uri $item.Url -Headers @{ 'User-Agent' = 'cli-buddy-starter' }
            foreach ($entry in $resp) {
                $relPath = if ($item.Rel) { Join-Path $item.Rel $entry.name } else { $entry.name }
                if ($entry.type -eq 'dir') {
                    New-Item -ItemType Directory -Path (Join-Path $target $relPath) -Force | Out-Null
                    $stack.Push(@{ Url = $entry.url; Rel = $relPath })
                } elseif ($entry.type -eq 'file') {
                    $destFile = Join-Path $target $relPath
                    Invoke-WebRequest -Uri $entry.download_url -OutFile $destFile -UseBasicParsing | Out-Null
                }
            }
        }
        $fileCount = (Get-ChildItem $target -Recurse -File | Measure-Object).Count
        Ok "$Skill - $fileCount files -> $target"
        return $true
    }

    $okCount = 0; $failCount = 0
    foreach ($s in $EnterpriseSkills) {
        try { if (Download-SkillFolder -Skill $s) { $okCount++ } }
        catch {
            Fail2 "$s failed: $($_.Exception.Message)"; $failCount++
            AddFailure 'Enterprise skills' $s $_.Exception.Message `
                "download install.ps1 and re-run: .\install.ps1 -Skip anthropic,community,plugins,snippet -Force"
        }
    }
    $summary['Enterprise skills'] = "$okCount ok, $failCount failed"
}

# ------------- 2. Anthropic doc skills -------------
Section 2 "Anthropic document skills (gh skill install)"

if (InSkip 'anthropic') {
    Warn2 "Skipping Anthropic skills (per -Skip)"
} elseif (-not $hasGh) {
    Warn2 "gh CLI not installed -> skipping Anthropic skills."
    Warn2 "Install from https://cli.github.com/ and re-run, or run:"
    Warn2 "  .\install.ps1 -Skip enterprise,plugins"
    $summary['Anthropic skills'] = 'skipped (gh not installed)'
} elseif (-not $ghAuthed) {
    Warn2 "gh CLI not authenticated -> skipping Anthropic skills."
    Warn2 "Run: gh auth login --hostname github.com"
    Warn2 "Then re-run this installer."
    $summary['Anthropic skills'] = 'skipped (gh not authenticated)'
} else {
    # Disable any interactive prompts from gh for the duration of this script.
    # Combined with --agent, this forces gh skill install to run fully non-interactively.
    # We do NOT capture gh's output (2>&1 makes stdout a pipe, which on some machines
    # causes gh to hang on a warning prompt we never see). Letting gh write to the
    # terminal gives the user real progress and preserves exit codes.
    $prevPromptDisabled = $env:GH_PROMPT_DISABLED
    $env:GH_PROMPT_DISABLED = '1'

    $okCount = 0; $failCount = 0
    foreach ($s in $AnthropicSkills) {
        Info "Installing anthropics/skills :: $s$(Desc $s)"
        try {
            & gh skill install anthropics/skills $s --agent github-copilot --scope user --force | Out-Host
            if ($LASTEXITCODE -eq 0) { Ok "$s installed"; $okCount++ }
            else {
                Fail2 "$s failed (exit $LASTEXITCODE) - see gh output above"; $failCount++
                AddFailure 'Anthropic skills' $s "gh exit $LASTEXITCODE" "gh skill install anthropics/skills $s --agent github-copilot --scope user --force"
            }
        } catch {
            Fail2 "$s failed: $($_.Exception.Message)"; $failCount++
            AddFailure 'Anthropic skills' $s $_.Exception.Message "gh skill install anthropics/skills $s --agent github-copilot --scope user --force"
        }
    }
    $summary['Anthropic skills'] = "$okCount ok, $failCount failed"
    $env:GH_PROMPT_DISABLED = $prevPromptDisabled
}

# ------------- 2b. Community skills -------------
Section '2b' "Community skills (gh skill install)"

if (InSkip 'community') {
    Warn2 "Skipping community skills (per -Skip)"
    $summary['Community skills'] = 'skipped (per -Skip)'
} elseif (-not $hasGh -or -not $ghAuthed) {
    Warn2 "gh CLI missing or not authenticated -> skipping community skills."
    $summary['Community skills'] = 'skipped (gh not ready)'
} else {
    $prevPromptDisabled = $env:GH_PROMPT_DISABLED
    $env:GH_PROMPT_DISABLED = '1'

    $okCount = 0; $failCount = 0

    foreach ($s in $SentrySkills) {
        Info "Installing Sentry01/copilot-cli-skills :: $s$(Desc $s)"
        try {
            & gh skill install Sentry01/copilot-cli-skills $s --agent github-copilot --scope user --force | Out-Host
            if ($LASTEXITCODE -eq 0) { Ok "$s installed"; $okCount++ }
            else {
                Fail2 "$s failed (exit $LASTEXITCODE) - see gh output above"; $failCount++
                AddFailure 'Community skills' $s "gh exit $LASTEXITCODE" "gh skill install Sentry01/copilot-cli-skills $s --agent github-copilot --scope user --force"
            }
        } catch {
            Fail2 "$s failed: $($_.Exception.Message)"; $failCount++
            AddFailure 'Community skills' $s $_.Exception.Message "gh skill install Sentry01/copilot-cli-skills $s --agent github-copilot --scope user --force"
        }
    }

    foreach ($s in $JimbanachSkills) {
        Info "Installing jimbanach/copilot-cli-starter :: $s@$JimbanachRef$(Desc $s)"
        try {
            & gh skill install jimbanach/copilot-cli-starter "$s@$JimbanachRef" --agent github-copilot --scope user --force | Out-Host
            if ($LASTEXITCODE -eq 0) { Ok "$s installed"; $okCount++ }
            else {
                Fail2 "$s failed (exit $LASTEXITCODE) - see gh output above"; $failCount++
                AddFailure 'Community skills' $s "gh exit $LASTEXITCODE" "gh skill install jimbanach/copilot-cli-starter $s@$JimbanachRef --agent github-copilot --scope user --force"
            }
        } catch {
            Fail2 "$s failed: $($_.Exception.Message)"; $failCount++
            AddFailure 'Community skills' $s $_.Exception.Message "gh skill install jimbanach/copilot-cli-starter $s@$JimbanachRef --agent github-copilot --scope user --force"
        }
    }

    $summary['Community skills'] = "$okCount ok, $failCount failed"
    $env:GH_PROMPT_DISABLED = $prevPromptDisabled
}

# ------------- 3. Copilot plugins -------------
Section 3 "Copilot CLI plugins"

if (InSkip 'plugins') {
    Warn2 "Skipping plugins (per -Skip)"
} elseif (-not $hasCopilot) {
    Warn2 "copilot CLI not found on PATH -> skipping plugins."
    Warn2 "Install Copilot CLI and re-run, or:"
    Warn2 "  .\install.ps1 -Skip plugins"
    $summary['Copilot plugins'] = 'skipped (copilot CLI not found)'
} else {
    $okCount = 0; $failCount = 0
    foreach ($p in $Plugins) {
        Info "Installing plugin :: $p$(Desc $p)"
        try {
            $out = & copilot plugin install $p 2>&1
            if ($LASTEXITCODE -eq 0) { Ok "$p installed"; $okCount++ }
            else {
                Fail2 "$p failed: $out"; $failCount++
                AddFailure 'Copilot plugins' $p $out "copilot plugin install $p"
            }
        } catch {
            Fail2 "$p failed: $($_.Exception.Message)"; $failCount++
            AddFailure 'Copilot plugins' $p $_.Exception.Message "copilot plugin install $p"
        }
    }
    $summary['Copilot plugins'] = "$okCount ok, $failCount failed"
    if ($okCount -gt 0) { Info "Restart any running Copilot CLI sessions for plugins to load." }
}

# ------------- 4. Instructions snippet (interactive or -AddSnippet) -------------
Section 4 "Instructions snippet (optional)"

if (InSkip 'snippet') {
    Warn2 "Skipping snippet step (per -Skip)"
    $summary['Instructions snippet'] = 'skipped (per -Skip)'
} else {
    $snippetUrl  = "https://raw.githubusercontent.com/$Repo/$Branch/copilot-instructions.snippet.md"
    $targetDir   = Join-Path $env:USERPROFILE '.copilot'
    $targetFile  = Join-Path $targetDir 'copilot-instructions.md'
    $beginMarker = '<!-- BEGIN: cli-buddy-starter enterprise skills v1 -->'
    $endMarker   = '<!-- END: cli-buddy-starter enterprise skills v1 -->'

    $fileExists = Test-Path $targetFile
    $blockPresent = $false
    $existingContent = ''
    if ($fileExists) {
        $existingContent = Get-Content $targetFile -Raw -ErrorAction SilentlyContinue
        if ($existingContent -and $existingContent.Contains($beginMarker)) { $blockPresent = $true }
    }

    Info "The snippet is a guidance block that makes Copilot CLI route Office files"
    Info "more predictably. Optional - the skills work without it."
    Info ""

    $proceed = $false
    if ($AddSnippet) {
        Info "Proceeding automatically (-AddSnippet)"
        $proceed = $true
    } else {
        try {
            $answer = Read-Host "Install/update snippet in $targetFile ? [y/N]"
            $proceed = ($answer -match '^(y|yes)$')
        } catch {
            Warn2 "Non-interactive session - skipping automatic install."
            $proceed = $false
        }
    }

    if ($proceed) {
        try {
            $snippet = (Invoke-WebRequest -Uri $snippetUrl -UseBasicParsing).Content
            if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }

            if ($blockPresent) {
                $bi = $existingContent.IndexOf($beginMarker)
                $ei = $existingContent.IndexOf($endMarker) + $endMarker.Length
                $newContent = $existingContent.Substring(0, $bi) + $snippet.TrimEnd() + $existingContent.Substring($ei)
                Set-Content -Path $targetFile -Value $newContent -Encoding UTF8
                Ok "Replaced snippet block in $targetFile"
                $summary['Instructions snippet'] = "updated ($targetFile)"
            } else {
                $needsSeparator = $fileExists -and $existingContent -and $existingContent.TrimEnd().Length -gt 0
                $toWrite = if ($needsSeparator) { "`r`n`r`n" + $snippet.TrimEnd() + "`r`n" } else { $snippet.TrimEnd() + "`r`n" }
                Add-Content -Path $targetFile -Value $toWrite -Encoding UTF8 -NoNewline
                Ok "Appended snippet to $targetFile"
                $summary['Instructions snippet'] = "installed ($targetFile)"
            }
            Info "To remove later: delete everything between the BEGIN / END markers."
        } catch {
            Fail2 "Snippet install failed: $($_.Exception.Message)"
            AddFailure 'Instructions snippet' 'append/replace' $_.Exception.Message "See README Step 4 for manual copy-paste"
            $summary['Instructions snippet'] = 'failed (see manual steps)'
            $proceed = $false
        }
    }

    if (-not $proceed -and $summary['Instructions snippet'] -notmatch 'installed|updated|failed|skipped') {
        Info ""
        Info "Manual steps:"
        Info "  1. Open: https://github.com/$Repo/blob/$Branch/copilot-instructions.snippet.md"
        Info "  2. Click Raw, copy EVERYTHING (including BEGIN/END markers)"
        Info "  3. Paste at the END of $targetFile (create if missing)"
        Info "  4. Do NOT replace anything above the block"
        Info ""
        Info "  (For a single repo only, paste into <repo>/.github/copilot-instructions.md instead.)"
        $summary['Instructions snippet'] = 'declined (see manual steps above)'
    }
}

# ------------- summary -------------
Write-Host ""
Write-Host "== Summary" -ForegroundColor White
foreach ($k in $summary.Keys) {
    "{0,-22} : {1}" -f $k, $summary[$k] | Write-Host
}

if ($failures.Count -gt 0) {
    Write-Host ""
    Write-Host "== Next steps for failures" -ForegroundColor Yellow
    Write-Host "The items below did not install. Most are fixable by re-running the specific command." -ForegroundColor DarkGray
    Write-Host ""
    foreach ($f in $failures) {
        Write-Host ("  [{0}] {1}" -f $f.Section, $f.Item) -ForegroundColor Yellow
        if ($f.Reason) {
            $line = ($f.Reason -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -First 1)
            if ($line) {
                if ($line.Length -gt 220) { $line = $line.Substring(0, 220) + '...' }
                Write-Host ("    reason : {0}" -f $line) -ForegroundColor DarkGray
            }
        }
        Write-Host ("    retry  : {0}" -f $f.Retry) -ForegroundColor DarkGray
    }
    Write-Host ""
    Write-Host "Common causes:" -ForegroundColor DarkGray
    Write-Host "  - git or gh CLI not installed  -> winget install Git.Git / winget install GitHub.cli" -ForegroundColor DarkGray
    Write-Host "  - gh CLI not authenticated     -> gh auth login --hostname github.com" -ForegroundColor DarkGray
    Write-Host "  - 'running scripts is disabled'-> Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned" -ForegroundColor DarkGray
    Write-Host "  - wrong plugin marketplace     -> check README troubleshooting section" -ForegroundColor DarkGray
    Write-Host "  - GitHub API rate limit        -> wait 60 min or authenticate gh CLI" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "Done. Open a new Copilot CLI session to pick up all changes." -ForegroundColor Green
Write-Host "Full docs: https://github.com/$Repo" -ForegroundColor DarkGray
Write-Host ""
