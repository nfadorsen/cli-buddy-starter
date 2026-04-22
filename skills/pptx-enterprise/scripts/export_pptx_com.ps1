[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$Path,

    [Parameter(Mandatory=$true)]
    [string]$OutDir
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $Path)) {
    Write-Error "File not found: $Path"
    exit 2
}

$deckPath = (Resolve-Path -LiteralPath $Path).Path
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$slidesDir = Join-Path $OutDir 'slides'
New-Item -ItemType Directory -Force -Path $slidesDir | Out-Null

$startedByUs = $false
$pptProcBefore = Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id
if (-not $pptProcBefore) { $startedByUs = $true }

$ppt = $null
$pres = $null

function Release-Com($o) {
    if ($o) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null } catch {}
    }
}

function Get-ShapeText($shape) {
    try {
        if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
            return $shape.TextFrame.TextRange.Text
        }
    } catch {}
    return ''
}

function Get-SlideTitle($slide) {
    try {
        foreach ($sh in $slide.Shapes) {
            try {
                if ($sh.Type -eq 14 -or $sh.HasTextFrame -eq -1) {
                    # PlaceholderType: ppPlaceholderTitle=13, ppPlaceholderCenterTitle=15
                    if ($sh.Type -eq 14) {
                        $ptype = $sh.PlaceholderFormat.Type
                        if ($ptype -eq 13 -or $ptype -eq 15) {
                            $t = Get-ShapeText $sh
                            if ($t) { return $t.Trim() }
                        }
                    }
                }
            } catch {}
        }
        # fallback: first non-empty text
        foreach ($sh in $slide.Shapes) {
            $t = Get-ShapeText $sh
            if ($t -and $t.Trim().Length -gt 0) { return $t.Trim() }
        }
    } catch {}
    return ''
}

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    # PowerPoint.Application cannot be fully hidden, but WindowState can be minimized.
    try { $ppt.WindowState = 2 } catch {}

    # Open ReadOnly, do not add to MRU, do not open with window
    # Open(FileName, ReadOnly, Untitled, WithWindow)
    $pres = $ppt.Presentations.Open($deckPath, $true, $false, $false)

    $slideWidth  = [double]$pres.PageSetup.SlideWidth
    $slideHeight = [double]$pres.PageSetup.SlideHeight
    $exportWidth = 1920
    if ($slideWidth -gt 0) {
        $exportHeight = [int]([math]::Round($exportWidth * ($slideHeight / $slideWidth)))
    } else {
        $exportHeight = 1080
    }

    $slides = @()
    $i = 0
    foreach ($slide in $pres.Slides) {
        $i++
        $allTextParts = @()
        $shapeTextCount = 0
        foreach ($sh in $slide.Shapes) {
            $t = Get-ShapeText $sh
            if ($t -and $t.Trim().Length -gt 0) {
                $allTextParts += $t.Trim()
                $shapeTextCount++
            }
        }
        $title = Get-SlideTitle $slide

        $notesText = ''
        try {
            if ($slide.HasNotesPage -eq -1) {
                foreach ($sh in $slide.NotesPage.Shapes) {
                    try {
                        if ($sh.PlaceholderFormat.Type -eq 2) { # ppPlaceholderBody
                            $nt = Get-ShapeText $sh
                            if ($nt) { $notesText = $nt.Trim() }
                        }
                    } catch {}
                }
            }
        } catch {}

        $pngName = ("Slide-{0:000}.png" -f $i)
        $pngPath = Join-Path $slidesDir $pngName
        try {
            $slide.Export($pngPath, 'PNG', $exportWidth, $exportHeight)
        } catch {
            Write-Warning ("Failed to export slide {0} as PNG: {1}" -f $i, $_.Exception.Message)
        }

        $slides += [ordered]@{
            slideNumber    = $i
            title          = $title
            allText        = ($allTextParts -join "`n")
            notesText      = $notesText
            shapeTextCount = $shapeTextCount
            pngPath        = $pngPath
        }
    }

    $result = [ordered]@{
        sourcePath   = $deckPath
        slideCount   = $slides.Count
        slideWidth   = $slideWidth
        slideHeight  = $slideHeight
        exportWidth  = $exportWidth
        exportHeight = $exportHeight
        slides       = $slides
    }

    $jsonPath = Join-Path $OutDir 'slides.json'
    $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonPath -Encoding UTF8

    Write-Host ("Exported {0} slides to {1}" -f $slides.Count, $OutDir)
    Write-Host ("slides.json: {0}" -f $jsonPath)
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
