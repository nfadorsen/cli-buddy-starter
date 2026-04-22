<#
.SYNOPSIS
    Export every worksheet in an Excel workbook to CSV via Excel COM.
    Designed for IRM / sensitivity-labeled workbooks that Python libraries
    cannot read. Does NOT modify the source workbook and does NOT touch
    sensitivity labels.

.PARAMETER Path
    Full path to the .xlsx / .xlsm / .xls workbook.

.PARAMETER OutDir
    Output directory. Defaults to "<workbook-folder>\exports"; if that folder
    cannot be created, falls back to "<cwd>\exports".

.OUTPUTS
    <OutDir>\<sheet>.csv           - one CSV per worksheet
    <OutDir>\sheets.json           - manifest [{name,rows,cols,csv}]

.NOTES
    ReadOnly open. Each sheet is copied to a temporary workbook before SaveAs
    CSV (xlCSV=6), then that temp workbook is closed without saving. The
    source is never written.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$OutDir
)

$ErrorActionPreference = 'Stop'

function Resolve-OutDir {
    param([string]$WorkbookPath, [string]$Override)
    if ($Override) {
        $null = New-Item -ItemType Directory -Force -Path $Override
        return (Resolve-Path $Override).Path
    }
    $preferred = Join-Path (Split-Path -Parent $WorkbookPath) 'exports'
    try {
        $null = New-Item -ItemType Directory -Force -Path $preferred -ErrorAction Stop
        # Writability probe
        $probe = Join-Path $preferred ('.probe_' + [guid]::NewGuid().ToString('N') + '.tmp')
        Set-Content -LiteralPath $probe -Value 'ok' -ErrorAction Stop
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        return (Resolve-Path $preferred).Path
    } catch {
        $fallback = Join-Path (Get-Location).Path 'exports'
        $null = New-Item -ItemType Directory -Force -Path $fallback
        Write-Warning "Workbook folder not writable; using $fallback"
        return (Resolve-Path $fallback).Path
    }
}

function Get-SafeName {
    param([string]$Name)
    $invalid = [IO.Path]::GetInvalidFileNameChars() + @('[', ']', '#', '%', '&', '{', '}')
    $safe = $Name
    foreach ($c in $invalid) { $safe = $safe -replace ([regex]::Escape([string]$c)), '_' }
    $safe = $safe.Trim()
    if (-not $safe) { $safe = 'Sheet' }
    return $safe
}

function Release-Com {
    param($obj)
    if ($null -ne $obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch {}
    }
}

if (-not (Test-Path -LiteralPath $Path)) {
    throw "File not found: $Path"
}
$Path = (Resolve-Path -LiteralPath $Path).Path
$OutDir = Resolve-OutDir -WorkbookPath $Path -Override $OutDir

Write-Host "Workbook: $Path"
Write-Host "Output  : $OutDir"

$xlCSV = 6
$excel = $null
$wb = $null
$manifest = @()

try {
    # Start Excel with retries; COM startup can race with other Office processes.
    $tries = 0
    while (-not $excel -and $tries -lt 5) {
        try {
            $excel = New-Object -ComObject Excel.Application
        } catch {
            $tries++
            Start-Sleep -Seconds 5
        }
    }
    if (-not $excel) { throw "Unable to start Excel.Application after $tries tries." }

    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    # Visible stays $false; some labeled files dislike hidden + protected view, but ReadOnly+no-alerts covers it.

    # Open ReadOnly. 3rd positional arg = ReadOnly.
    $wb = $excel.Workbooks.Open($Path, 0, $true)

    $sheetCount = $wb.Worksheets.Count
    Write-Host "Sheets  : $sheetCount"

    $usedNames = @{}
    for ($i = 1; $i -le $sheetCount; $i++) {
        $ws = $wb.Worksheets.Item($i)
        $origName = [string]$ws.Name
        $safe = Get-SafeName -Name $origName

        # Ensure uniqueness on disk
        $candidate = $safe
        $suffix = 2
        while ($usedNames.ContainsKey($candidate.ToLower())) {
            $candidate = "{0}_{1}" -f $safe, $suffix
            $suffix++
        }
        $usedNames[$candidate.ToLower()] = $true
        $csvPath = Join-Path $OutDir ($candidate + '.csv')

        $used = $ws.UsedRange
        $rows = [int]$used.Rows.Count
        $cols = [int]$used.Columns.Count

        Write-Host ("  [{0}/{1}] {2}  ({3} x {4})  -> {5}" -f $i, $sheetCount, $origName, $rows, $cols, (Split-Path -Leaf $csvPath))

        # Copy sheet to a new standalone workbook so SaveAs-CSV only affects that copy.
        $tempWb = $null
        try {
            $ws.Copy()                # with no args, creates a new workbook with the single sheet copy
            $tempWb = $excel.ActiveWorkbook
            if ($tempWb -eq $null) { throw "Worksheet.Copy() did not produce an ActiveWorkbook for sheet '$origName'." }
            # SaveAs CSV (xlCSV = 6)
            $tempWb.SaveAs($csvPath, $xlCSV)
            $tempWb.Close($false)     # do not save again
        } finally {
            Release-Com $tempWb
        }

        $manifest += [ordered]@{
            name = $origName
            rows = $rows
            cols = $cols
            csv  = $csvPath
        }

        Release-Com $used
        Release-Com $ws
    }

    # Write manifest
    $manifestPath = Join-Path $OutDir 'sheets.json'
    ($manifest | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    Write-Host "Manifest: $manifestPath"
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {}; Release-Com $wb }
    if ($excel) {
        try { $excel.Quit() } catch {}
        Release-Com $excel
    }
    [GC]::Collect() | Out-Null
    [GC]::WaitForPendingFinalizers() | Out-Null
    [GC]::Collect() | Out-Null
}

Write-Host "Done."
