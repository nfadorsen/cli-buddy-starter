[CmdletBinding()]
<#
.SYNOPSIS
  Preflight health check for Office COM automation on this machine.

.DESCRIPTION
  Verifies the preconditions required to read/edit PowerPoint decks that
  contain Office waterfall charts (chart type 119) or other chart objects
  whose underlying data lives in an embedded Excel workbook.

  This script is deterministic and fail-fast. It is NOT a repair tool:
  if anything fails it prints a remediation checklist and exits non-zero.
  Callers must NOT attempt screenshot/image/shape-drawing workarounds.

.PARAMETER DeckPath
  Optional. Path to a .pptx file. When supplied:
    - Detects OLE2/IRM header (D0 CF 11 E0) vs zip header (50 4B).
    - Opens the deck read-only in PowerPoint COM.
    - If -RequireWaterfall is also set, finds the first chart whose
      ChartType is 119 and tries Chart.ChartData.Activate() on it.

.PARAMETER RequireWaterfall
  If set, the script will additionally locate a waterfall chart
  (ChartType 119) in the supplied deck and attempt to activate its
  embedded workbook via Excel COM. Failure => abort.

.PARAMETER Json
  Emit a single-line JSON result object instead of human-readable text.

.OUTPUTS
  Exit codes:
    0  All required checks passed.
    2  Bad arguments or file-not-found.
    3  PowerPoint COM could not instantiate or open the deck.
    5  Waterfall required but ChartData.Activate() failed (the
       0x80080005 environmental failure surfaces here when the
       PowerPoint-spawned Excel backend can't start).
    6  Waterfall required but no waterfall chart found in the deck.

.NOTES
  Local-only. No network. No repair. No fallbacks.

  Design notes (learned the hard way):
  - On this machine, cold New-Object -ComObject Excel.Application
    triggers sign-in dialogs that fail with 0x80080005 and leave
    headless EXCEL.EXE orphans behind. This script therefore does
    NOT instantiate Excel explicitly. Excel is spawned by PowerPoint
    internally when Chart.ChartData.Activate() is called, which is
    the reliable path.
  - Every COM session that touches charts leaves at least one
    headless EXCEL.EXE orphan behind. This script sweeps such
    orphans (processes with empty MainWindowTitle) before starting,
    and also force-kills any Office PID it spawned on exit.
#>
param(
    [string]$DeckPath,
    [switch]$RequireWaterfall,
    [switch]$Json
)

$ErrorActionPreference = 'Stop'

# Sweep headless Office orphans first. These are zombies from prior COM
# sessions (no MainWindowTitle) and block fresh COM activations with
# 0x80080005. Only kill processes with no visible window — leave the
# user's interactive Excel/PowerPoint windows alone.
$orphansKilled = @()
Get-Process -Name EXCEL,POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { [string]::IsNullOrEmpty($_.MainWindowTitle) } |
    ForEach-Object {
        try { Stop-Process -Id $_.Id -Force -ErrorAction Stop; $orphansKilled += "$($_.Name):$($_.Id)" } catch {}
    }
if ($orphansKilled.Count -gt 0) { Start-Sleep -Seconds 2 }

# Remember which Office PIDs existed before so we can force-kill any we spawned on exit.
$prePids = Get-Process -Name EXCEL,POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }

$result = [ordered]@{
    timestamp         = (Get-Date).ToString('o')
    deckPath          = $DeckPath
    requireWaterfall  = [bool]$RequireWaterfall
    headerHex         = $null
    containerType     = $null
    orphansKilled     = $orphansKilled
    powerPointCom     = $false
    powerPointOpen    = $false
    waterfallFound    = $false
    waterfallActivate = $false
    errors            = @()
    remediation       = @()
    exitCode          = 0
}

function Fail([int]$code, [string]$message) {
    $result.errors += $message
    $result.exitCode = $code
    if ($result.remediation.Count -eq 0) {
        $result.remediation = @(
            'Close all open Excel and PowerPoint windows.',
            'End any orphaned EXCEL.EXE and POWERPNT.EXE processes in Task Manager.',
            'Open Excel once interactively (let it finish any sign-in), then close it.',
            'Re-run this preflight. If it still fails, do NOT proceed with the chart edit.'
        )
    }
    Emit
    exit $code
}

function Emit {
    if ($Json) {
        $result | ConvertTo-Json -Compress -Depth 5
        return
    }
    Write-Host ''
    Write-Host '--- Office Automation Preflight ---'
    Write-Host ("DeckPath           : {0}" -f ($result.deckPath   | ForEach-Object { if ($_) { $_ } else { '(none)' } }))
    Write-Host ("ContainerType      : {0}" -f ($result.containerType | ForEach-Object { if ($_) { $_ } else { '(n/a)' } }))
    Write-Host ("HeaderHex          : {0}" -f ($result.headerHex  | ForEach-Object { if ($_) { $_ } else { '(n/a)' } }))
    Write-Host ("PowerPoint COM     : {0}" -f $result.powerPointCom)
    Write-Host ("PowerPoint Open    : {0}" -f $result.powerPointOpen)
    Write-Host ("Orphans Swept      : {0}" -f (($result.orphansKilled -join ', ') | ForEach-Object { if ($_) { $_ } else { '(none)' } }))
    Write-Host ("Waterfall Found    : {0}" -f $result.waterfallFound)
    Write-Host ("Waterfall Activate : {0}" -f $result.waterfallActivate)
    Write-Host ("Exit Code          : {0}" -f $result.exitCode)
    if ($result.errors.Count -gt 0) {
        Write-Host ''
        Write-Host 'Errors:' -ForegroundColor Yellow
        foreach ($e in $result.errors) { Write-Host "  - $e" -ForegroundColor Yellow }
    }
    if ($result.remediation.Count -gt 0) {
        Write-Host ''
        Write-Host 'Remediation:' -ForegroundColor Cyan
        foreach ($r in $result.remediation) { Write-Host "  - $r" -ForegroundColor Cyan }
    }
}

function Release-Com($o) {
    if ($o) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null } catch {}
    }
}

# 1. Validate DeckPath (if supplied) and sniff header
if ($DeckPath) {
    if (-not (Test-Path -LiteralPath $DeckPath)) {
        Fail 2 "File not found: $DeckPath"
    }
    $resolved = (Resolve-Path -LiteralPath $DeckPath).Path
    $result.deckPath = $resolved
    $bytes = New-Object byte[] 8
    $fs = [System.IO.File]::Open($resolved, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try { [void]$fs.Read($bytes, 0, 8) } finally { $fs.Close(); $fs.Dispose() }
    $result.headerHex = ($bytes | ForEach-Object { $_.ToString('X2') }) -join ' '
    if ($bytes[0] -eq 0x50 -and $bytes[1] -eq 0x4B) {
        $result.containerType = 'zip'
    } elseif ($bytes[0] -eq 0xD0 -and $bytes[1] -eq 0xCF -and $bytes[2] -eq 0x11 -and $bytes[3] -eq 0xE0) {
        $result.containerType = 'ole2'
    } else {
        $result.containerType = 'unknown'
    }
} elseif ($RequireWaterfall) {
    Fail 2 '-RequireWaterfall requires -DeckPath.'
}

# 2. PowerPoint COM
$ppt = $null; $pres = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $result.powerPointCom = $true
} catch {
    Fail 3 ("PowerPoint COM could not instantiate: " + $_.Exception.Message)
}

try {
    # NOTE: We deliberately do NOT instantiate Excel.Application here.
    # On this machine, cold New-Object -ComObject Excel.Application
    # reliably triggers sign-in dialogs that fail with 0x80080005 and
    # leave headless orphans behind. Instead, we let PowerPoint spawn
    # Excel internally via Chart.ChartData.Activate() in the waterfall
    # step below. That path is reliable once orphans are cleared.

    if ($DeckPath) {
        # ReadOnly open so we never mutate the file during preflight
        $pres = $ppt.Presentations.Open($result.deckPath, $true, $false, $false)
        $result.powerPointOpen = $true
    }

    # 4. Waterfall check (only if requested and deck opened)
    if ($RequireWaterfall -and $pres) {
        $targetChart = $null
        foreach ($slide in $pres.Slides) {
            foreach ($sh in $slide.Shapes) {
                $hasChart = $false
                try { $hasChart = ($sh.HasChart -eq -1) } catch {}
                if ($hasChart) {
                    try {
                        if ($sh.Chart.ChartType -eq 119) {
                            $targetChart = $sh.Chart
                            break
                        }
                    } catch {}
                }
            }
            if ($targetChart) { break }
        }

        if (-not $targetChart) {
            Fail 6 'No waterfall chart (ChartType=119) found in deck; cannot verify chart-data activation.'
        }
        $result.waterfallFound = $true

        try {
            $targetChart.ChartData.Activate()
            Start-Sleep -Milliseconds 500
            $wb = $targetChart.ChartData.Workbook
            if (-not $wb) {
                Fail 5 'Chart.ChartData.Activate() returned but Workbook is null. Excel cannot serve the embedded workbook.'
            }
            $ws = $wb.Worksheets.Item(1)
            $rows = $ws.UsedRange.Rows.Count
            if ($rows -lt 1) {
                Fail 5 'Waterfall workbook opened but no rows present; cannot safely edit.'
            }
            $result.waterfallActivate = $true
            try { $wb.Close($false) } catch {}
        } catch {
            Fail 5 ("Chart.ChartData.Activate() failed: " + $_.Exception.Message)
        }
    }
}
finally {
    if ($pres) {
        try { $pres.Close() } catch {}
        Release-Com $pres
    }
    if ($ppt) {
        try { $ppt.Quit() } catch {}
        Release-Com $ppt
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    Start-Sleep -Seconds 2

    # Force-kill any Office PIDs we spawned (post minus pre).
    $postPids = Get-Process -Name EXCEL,POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }
    $spawned = $postPids | Where-Object { $prePids -notcontains $_ }
    foreach ($p in $spawned) {
        try { Stop-Process -Id $p -Force -ErrorAction Stop } catch {}
    }
}

$result.exitCode = 0
Emit
exit 0
