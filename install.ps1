# CLI Buddy Starter - installer
# Downloads 3 enterprise Copilot CLI skills into ~/.copilot/skills/
#   - pptx-enterprise
#   - docx-enterprise
#   - excel-enterprise
#
# What this script does:
#   - Creates ~/.copilot/skills/ if missing
#   - Downloads each skill folder from this GitHub repo via the GitHub API
#   - Writes files under ~/.copilot/skills/<skill-name>/
#   - Prints a summary of what was installed and where
#
# What this script does NOT do:
#   - No admin / elevation
#   - No registry changes, no scheduled tasks, no services
#   - No telemetry, no network beyond GitHub API / raw.githubusercontent.com
#   - No changes to Office, Windows, or any other installed application
#   - No changes to your Copilot CLI authentication
#
# To uninstall: delete the 3 folders under ~/.copilot/skills/.

[CmdletBinding()]
param(
    [string]$Repo   = 'nfadorsen/cli-buddy-starter',
    [string]$Branch = 'main',
    [string[]]$Skills = @('pptx-enterprise','docx-enterprise','excel-enterprise'),
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

function Write-Info($msg)    { Write-Host "  $msg" -ForegroundColor Cyan }
function Write-Ok($msg)      { Write-Host "  OK  $msg" -ForegroundColor Green }
function Write-Warn2($msg)   { Write-Host "  !!  $msg" -ForegroundColor Yellow }

Write-Host ""
Write-Host "CLI Buddy Starter - installing enterprise skills" -ForegroundColor White
Write-Host "Source : https://github.com/$Repo (branch: $Branch)" -ForegroundColor DarkGray
Write-Host ""

$skillsRoot = Join-Path $env:USERPROFILE '.copilot\skills'
if (-not (Test-Path $skillsRoot)) {
    New-Item -ItemType Directory -Path $skillsRoot -Force | Out-Null
    Write-Info "Created $skillsRoot"
}

function Download-SkillFolder {
    param([string]$Skill)

    $apiUrl = "https://api.github.com/repos/$Repo/contents/skills/$Skill" + "?ref=$Branch"
    $target = Join-Path $skillsRoot $Skill

    if ((Test-Path $target) -and -not $Force) {
        Write-Warn2 "$Skill already exists at $target (use -Force to overwrite). Skipping."
        return
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
    Write-Ok "$Skill - $fileCount files -> $target"
}

foreach ($s in $Skills) {
    Write-Info "Downloading $s..."
    try {
        Download-SkillFolder -Skill $s
    } catch {
        Write-Warn2 "Failed to install $s : $($_.Exception.Message)"
    }
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host "Installed skills:" -ForegroundColor White
Get-ChildItem $skillsRoot -Directory |
    Where-Object { $Skills -contains $_.Name } |
    ForEach-Object { "  - $($_.FullName)" } |
    Write-Host

Write-Host ""
Write-Host "Next steps:" -ForegroundColor White
Write-Host "  1. Open a new Copilot CLI session. The skills will be auto-detected."
Write-Host "  2. Try a prompt like: 'Inspect <some>.pptx and tell me the structure.'"
Write-Host "  3. (Optional) Copy copilot-instructions.sample.md from this repo into"
Write-Host "     your own project's .github/copilot-instructions.md and edit to taste."
Write-Host ""
Write-Host "To uninstall: delete the folders listed above." -ForegroundColor DarkGray
Write-Host ""
