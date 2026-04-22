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

$docPath = (Resolve-Path -LiteralPath $Path).Path
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

# Track pre-existing WINWORD processes so we only kill ones we spawned.
$wordProcBefore = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$startedByUs = ($wordProcBefore.Count -eq 0)

$word = $null
$doc  = $null

function Release-Com($o) {
    if ($o) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null } catch {}
    }
}

function Safe-Text($s) {
    if ($null -eq $s) { return '' }
    # Word often returns text with a trailing BEL (0x07) for end-of-cell markers; strip control chars.
    return ($s -replace "[\x00-\x08\x0B\x0C\x0E-\x1F]", '').TrimEnd([char]13, [char]7)
}

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0  # wdAlertsNone

    # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
    $doc = $word.Documents.Open($docPath, $false, $true, $false)

    # Structural counts
    $pageCount = 0
    try { $pageCount = [int]$doc.ComputeStatistics(2) } catch {}  # wdStatisticPages=2
    $wordCount = 0
    try { $wordCount = [int]$doc.ComputeStatistics(0) } catch {}  # wdStatisticWords=0

    # Sensitivity label (best-effort; available on Word 2019+ in M365 tenant)
    $labelName = ''
    $labelId   = ''
    try {
        $label = $doc.SensitivityLabel.GetLabel()
        if ($label) {
            $labelName = [string]$label.LabelName
            $labelId   = [string]$label.LabelId
        }
    } catch {}

    # Paragraphs with style + outline level
    $paragraphs = New-Object System.Collections.ArrayList
    $plainLines = New-Object System.Collections.ArrayList
    $pIndex = 0
    foreach ($p in $doc.Paragraphs) {
        $pIndex++
        $text = Safe-Text $p.Range.Text
        $style = ''
        try { $style = [string]$p.Style.NameLocal } catch {}
        $outline = 0
        try { $outline = [int]$p.OutlineLevel } catch {}

        [void]$paragraphs.Add([ordered]@{
            index        = $pIndex
            style        = $style
            outlineLevel = $outline
            text         = $text
        })
        if ($text.Length -gt 0) { [void]$plainLines.Add($text) }
    }

    # Tracked changes (revisions)
    $revisions = New-Object System.Collections.ArrayList
    try {
        $rIndex = 0
        foreach ($rev in $doc.Revisions) {
            $rIndex++
            $author = ''
            try { $author = [string]$rev.Author } catch {}
            $date = ''
            try { $date = ([DateTime]$rev.Date).ToString('o') } catch {}
            $type = 0
            try { $type = [int]$rev.Type } catch {}
            $revText = ''
            try { $revText = Safe-Text $rev.Range.Text } catch {}

            [void]$revisions.Add([ordered]@{
                index  = $rIndex
                author = $author
                date   = $date
                type   = $type    # 1=wdRevisionInsert, 2=wdRevisionDelete, etc.
                text   = $revText
            })
        }
    } catch {}

    # Comments
    $comments = New-Object System.Collections.ArrayList
    try {
        $cIndex = 0
        foreach ($cm in $doc.Comments) {
            $cIndex++
            $author = ''
            try { $author = [string]$cm.Author } catch {}
            $initials = ''
            try { $initials = [string]$cm.Initial } catch {}
            $date = ''
            try { $date = ([DateTime]$cm.Date).ToString('o') } catch {}
            $scope = ''
            try { $scope = Safe-Text $cm.Scope.Text } catch {}
            $cmText = ''
            try { $cmText = Safe-Text $cm.Range.Text } catch {}

            [void]$comments.Add([ordered]@{
                index    = $cIndex
                author   = $author
                initials = $initials
                date     = $date
                anchor   = $scope
                text     = $cmText
            })
        }
    } catch {}

    $result = [ordered]@{
        sourcePath     = $docPath
        pageCount      = $pageCount
        wordCount      = $wordCount
        paragraphCount = $paragraphs.Count
        revisionCount  = $revisions.Count
        commentCount   = $comments.Count
        sensitivityLabel = [ordered]@{
            name = $labelName
            id   = $labelId
        }
        paragraphs = $paragraphs
        revisions  = $revisions
        comments   = $comments
    }

    $jsonPath = Join-Path $OutDir 'document.json'
    $txtPath  = Join-Path $OutDir 'document.txt'

    $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
    ($plainLines -join [Environment]::NewLine) | Set-Content -LiteralPath $txtPath -Encoding UTF8

    Write-Host ("Exported {0} paragraphs, {1} revisions, {2} comments" -f $paragraphs.Count, $revisions.Count, $comments.Count)
    Write-Host ("document.json: {0}" -f $jsonPath)
    Write-Host ("document.txt:  {0}" -f $txtPath)
}
finally {
    if ($doc) {
        try { $doc.Close([ref]$false) } catch {}
        Release-Com $doc
    }
    if ($word) {
        try { $word.Quit() } catch {}
        Release-Com $word
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()

    if ($startedByUs) {
        Start-Sleep -Milliseconds 300
        Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object {
            if (-not ($wordProcBefore -contains $_.Id)) {
                try { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue } catch {}
            }
        }
    }
}
