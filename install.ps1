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
    [switch]$Force
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

Write-Host ""
Write-Host "CLI Buddy Starter - Copilot CLI setup" -ForegroundColor White
Write-Host "Source: https://github.com/$Repo (branch: $Branch)" -ForegroundColor DarkGray

# ------------- pre-flight -------------
Section 0 "Pre-flight checks"

$hasGh       = HasCommand gh
$hasCopilot  = HasCommand copilot
$ghAuthed    = $false
if ($hasGh) {
    try {
        gh auth status 2>&1 | Out-Null
        $ghAuthed = ($LASTEXITCODE -eq 0)
    } catch { $ghAuthed = $false }
}

if ($hasGh)      { Ok "gh CLI found" }            else { Warn2 "gh CLI not found (needed for Anthropic skills)" }
if ($ghAuthed)   { Ok "gh CLI authenticated" }    elseif ($hasGh) { Warn2 "gh CLI not authenticated (run: gh auth login)" }
if ($hasCopilot) { Ok "copilot CLI found" }       else { Warn2 "copilot CLI not found (needed for plugins)" }

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
        catch { Fail2 "$s failed: $($_.Exception.Message)"; $failCount++ }
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
    $okCount = 0; $failCount = 0
    foreach ($s in $AnthropicSkills) {
        Info "Installing anthropics/skills :: $s"
        try {
            $out = & gh skill install anthropics/skills $s --scope user --force 2>&1
            if ($LASTEXITCODE -eq 0) { Ok "$s installed"; $okCount++ }
            else { Fail2 "$s failed: $out"; $failCount++ }
        } catch {
            Fail2 "$s failed: $($_.Exception.Message)"; $failCount++
        }
    }
    $summary['Anthropic skills'] = "$okCount ok, $failCount failed"
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
    $okCount = 0; $failCount = 0

    foreach ($s in $SentrySkills) {
        Info "Installing Sentry01/copilot-cli-skills :: $s"
        try {
            $out = & gh skill install Sentry01/copilot-cli-skills $s --scope user --force 2>&1
            if ($LASTEXITCODE -eq 0) { Ok "$s installed"; $okCount++ }
            else { Fail2 "$s failed: $out"; $failCount++ }
        } catch {
            Fail2 "$s failed: $($_.Exception.Message)"; $failCount++
        }
    }

    foreach ($s in $JimbanachSkills) {
        Info "Installing jimbanach/copilot-cli-starter :: $s@$JimbanachRef"
        try {
            $out = & gh skill install jimbanach/copilot-cli-starter "$s@$JimbanachRef" --scope user --force 2>&1
            if ($LASTEXITCODE -eq 0) { Ok "$s installed"; $okCount++ }
            else { Fail2 "$s failed: $out"; $failCount++ }
        } catch {
            Fail2 "$s failed: $($_.Exception.Message)"; $failCount++
        }
    }

    $summary['Community skills'] = "$okCount ok, $failCount failed"
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
        Info "Installing plugin :: $p"
        try {
            $out = & copilot plugin install $p 2>&1
            if ($LASTEXITCODE -eq 0) { Ok "$p installed"; $okCount++ }
            else { Fail2 "$p failed: $out"; $failCount++ }
        } catch {
            Fail2 "$p failed: $($_.Exception.Message)"; $failCount++
        }
    }
    $summary['Copilot plugins'] = "$okCount ok, $failCount failed"
    if ($okCount -gt 0) { Info "Restart any running Copilot CLI sessions for plugins to load." }
}

# ------------- 4. Instructions snippet (manual) -------------
Section 4 "Instructions snippet (optional, manual)"

if (InSkip 'snippet') {
    Warn2 "Skipping snippet step (per -Skip)"
    $summary['Instructions snippet'] = 'skipped (per -Skip)'
} else {
    Info "The instructions snippet is a small guidance block that makes Copilot CLI"
    Info "route Office files more predictably. It's optional - the skills work without it."
    Info ""
    Info "To add it:"
    Info "  1. Open: https://github.com/$Repo/blob/$Branch/copilot-instructions.snippet.md"
    Info "  2. Click Raw, copy EVERYTHING (including BEGIN/END markers)"
    Info "  3. Paste at the END of your .github/copilot-instructions.md"
    Info "  4. Do NOT replace anything above the block"
}

# ------------- summary -------------
Write-Host ""
Write-Host "== Summary" -ForegroundColor White
foreach ($k in $summary.Keys) {
    "{0,-22} : {1}" -f $k, $summary[$k] | Write-Host
}

Write-Host ""
Write-Host "Done. Open a new Copilot CLI session to pick up all changes." -ForegroundColor Green
Write-Host "Full docs: https://github.com/$Repo" -ForegroundColor DarkGray
Write-Host ""
