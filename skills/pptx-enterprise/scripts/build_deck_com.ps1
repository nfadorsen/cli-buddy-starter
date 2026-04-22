[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$OutPath,

    [Parameter(Mandatory=$true)]
    [string]$SpecPath,

    [string]$BaseDeckPath
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $SpecPath)) {
    Write-Error "Spec not found: $SpecPath"
    exit 2
}

$outDir = Split-Path -Parent $OutPath
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
}

$spec = Get-Content -LiteralPath $SpecPath -Raw | ConvertFrom-Json

# Decide whether to use BaseDeckPath: only if zip-backed PPTX.
$useBase = $false
if ($BaseDeckPath) {
    if (Test-Path -LiteralPath $BaseDeckPath) {
        $bytes = New-Object byte[] 4
        $fs = [System.IO.File]::Open((Resolve-Path -LiteralPath $BaseDeckPath).Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try { [void]$fs.Read($bytes, 0, 4) } finally { $fs.Close(); $fs.Dispose() }
        if ($bytes[0] -eq 0x50 -and $bytes[1] -eq 0x4B) {
            $useBase = $true
        } else {
            Write-Warning "BaseDeckPath is not zip-backed PPTX; ignoring and creating a new clean deck."
        }
    } else {
        Write-Warning "BaseDeckPath not found; ignoring."
    }
}

$startedByUs = $false
$pptProcBefore = Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id
if (-not $pptProcBefore) { $startedByUs = $true }

function Release-Com($o) {
    if ($o) { try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null } catch {} }
}

$ppt = $null
$pres = $null

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    try { $ppt.WindowState = 2 } catch {}

    if ($useBase) {
        $pres = $ppt.Presentations.Open((Resolve-Path -LiteralPath $BaseDeckPath).Path, $false, $false, $false)
    } else {
        $pres = $ppt.Presentations.Add($false)
    }

    # Layout enums: ppLayoutTitle=1, ppLayoutText=2, ppLayoutTitleOnly=11, ppLayoutBlank=12
    function Map-Layout($name) {
        switch (($name | ForEach-Object { "$_".ToLowerInvariant() })) {
            'title'   { return 1 }
            'content' { return 2 }
            'blank'   { return 12 }
            default   { return 2 }
        }
    }

    function Set-Title($slide, [string]$text) {
        if (-not $text) { return }
        try {
            if ($slide.Shapes.HasTitle -eq -1) {
                $slide.Shapes.Title.TextFrame.TextRange.Text = $text
                return
            }
        } catch {}
        # Fallback: add a textbox at the top
        try {
            $tb = $slide.Shapes.AddTextbox(1, 36, 24, 888, 72)
            $tb.TextFrame.TextRange.Text = $text
            $tb.TextFrame.TextRange.Font.Size = 32
            $tb.TextFrame.TextRange.Font.Bold = -1
        } catch {}
    }

    function Set-Bullets($slide, [string[]]$bullets, [string]$subtitle) {
        $bodyShape = $null
        try {
            foreach ($sh in $slide.Shapes) {
                try {
                    if ($sh.Type -eq 14) {
                        $pt = $sh.PlaceholderFormat.Type
                        # body=2, subtitle=4, object=7, centerTitle=15, title=13
                        if ($pt -eq 2 -or $pt -eq 4 -or $pt -eq 7) { $bodyShape = $sh; break }
                    }
                } catch {}
            }
        } catch {}

        $joined = @()
        if ($subtitle) { $joined += $subtitle }
        if ($bullets) { $joined += $bullets }
        if ($joined.Count -eq 0) { return }
        $text = ($joined -join "`r")

        if ($bodyShape) {
            try {
                $bodyShape.TextFrame.TextRange.Text = $text
                return
            } catch {}
        }
        # Fallback textbox
        try {
            $tb = $slide.Shapes.AddTextbox(1, 36, 120, 888, 420)
            $tb.TextFrame.TextRange.Text = $text
            $tb.TextFrame.TextRange.Font.Size = 20
        } catch {}
    }

    function Set-Notes($slide, [string]$notes) {
        if (-not $notes) { return }
        try {
            foreach ($sh in $slide.NotesPage.Shapes) {
                try {
                    if ($sh.PlaceholderFormat.Type -eq 2) {
                        $sh.TextFrame.TextRange.Text = $notes
                        return
                    }
                } catch {}
            }
        } catch {}
    }

    $slideSpecs = @()
    if ($spec.slides) { $slideSpecs = @($spec.slides) }

    $existingCount = 0
    try { $existingCount = [int]$pres.Slides.Count } catch {}

    $idx = 0
    foreach ($s in $slideSpecs) {
        $idx++
        $layoutNum = Map-Layout $s.layout
        $insertAt = $existingCount + $idx
        $slide = $pres.Slides.Add($insertAt, $layoutNum)

        Set-Title   $slide $s.title
        Set-Bullets $slide @($s.bullets) $s.subtitle
        Set-Notes   $slide $s.notes
    }

    # ppSaveAsOpenXMLPresentation = 24
    $pres.SaveAs($OutPath, 24)
    Write-Host ("Saved deck to {0}" -f $OutPath)
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

    if ($startedByUs) {
        Start-Sleep -Milliseconds 300
        Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | ForEach-Object {
            if (-not ($pptProcBefore -contains $_.Id)) {
                try { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue } catch {}
            }
        }
    }
}
