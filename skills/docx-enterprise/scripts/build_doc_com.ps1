[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$Path,

    [Parameter(Mandatory=$true)]
    [ValidateSet('accept-changes','find-replace','add-comment','extract-redlines')]
    [string]$Mode,

    [string]$Find,
    [string]$Replace,
    [string]$Anchor,
    [string]$CommentText,
    [string]$Author = 'Reviewer',
    [string]$OutPath,
    [string]$OutDir
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $Path)) {
    Write-Error "File not found: $Path"
    exit 2
}

$srcPath = (Resolve-Path -LiteralPath $Path).Path

# If OutPath specified, operate on a copy so the original is never touched.
$workingPath = $srcPath
if ($OutPath) {
    $OutParent = Split-Path -Parent $OutPath
    if ($OutParent) { New-Item -ItemType Directory -Force -Path $OutParent | Out-Null }
    Copy-Item -LiteralPath $srcPath -Destination $OutPath -Force
    $workingPath = (Resolve-Path -LiteralPath $OutPath).Path
}

# Track pre-existing WINWORD processes
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
    return ($s -replace "[\x00-\x08\x0B\x0C\x0E-\x1F]", '').TrimEnd([char]13, [char]7)
}

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    # extract-redlines is the only read-only mode
    $readOnly = ($Mode -eq 'extract-redlines')

    $doc = $word.Documents.Open($workingPath, $false, $readOnly, $false)

    switch ($Mode) {

        'accept-changes' {
            $before = [int]$doc.Revisions.Count
            if ($before -gt 0) {
                $doc.Revisions.AcceptAll()
            }
            $doc.Save()
            Write-Host ("Accepted {0} revisions. Saved in place at {1}" -f $before, $workingPath)
        }

        'find-replace' {
            if ([string]::IsNullOrEmpty($Find)) {
                throw "-Find is required for find-replace mode"
            }
            if ($null -eq $Replace) { $Replace = '' }

            $replacedTotal = 0

            # Body + headers + footers
            $storyRanges = @()
            $storyRanges += $doc.Content
            try {
                foreach ($section in $doc.Sections) {
                    foreach ($hf in @($section.Headers, $section.Footers)) {
                        foreach ($item in $hf) {
                            if ($item.Exists) { $storyRanges += $item.Range }
                        }
                    }
                }
            } catch {}

            foreach ($range in $storyRanges) {
                $find = $range.Find
                $find.ClearFormatting()
                $find.Replacement.ClearFormatting()
                $find.Text = [string]$Find
                $find.Replacement.Text = [string]$Replace
                $find.Forward = $true
                $find.Wrap = 1      # wdFindContinue
                $find.Format = $false
                $find.MatchCase = $false
                $find.MatchWholeWord = $false
                $find.MatchWildcards = $false
                # Execute with Replace:=wdReplaceAll (2)
                [void]$find.Execute([ref][string]$Find, [ref]$false, [ref]$false, [ref]$false, [ref]$false, [ref]$false, [ref]$true, [ref]1, [ref]$false, [ref][string]$Replace, [ref]2)
            }

            $doc.Save()
            Write-Host ("Find/replace complete. Pattern: '{0}' -> '{1}'. Saved in place at {2}" -f $Find, $Replace, $workingPath)
        }

        'add-comment' {
            if ([string]::IsNullOrEmpty($Anchor)) { throw "-Anchor is required for add-comment mode" }
            if ([string]::IsNullOrEmpty($CommentText)) { throw "-CommentText is required for add-comment mode" }

            $find = $doc.Content.Find
            $find.ClearFormatting()
            $find.Text = [string]$Anchor
            $find.Forward = $true
            $find.Wrap = 0  # wdFindStop

            $found = $find.Execute()
            if (-not $found) {
                throw "Anchor text not found in document: '$Anchor'"
            }

            $anchorRange = $doc.Content.Duplicate
            $anchorRange.Start = $find.Parent.Start
            $anchorRange.End   = $find.Parent.End

            $newComment = $doc.Comments.Add($anchorRange, [string]$CommentText)
            try { $newComment.Author = [string]$Author } catch {}

            $doc.Save()
            Write-Host ("Added comment anchored at '{0}'. Saved in place at {1}" -f $Anchor, $workingPath)
        }

        'extract-redlines' {
            $outDirEffective = $OutDir
            if (-not $outDirEffective) {
                $outDirEffective = Join-Path (Split-Path -Parent $workingPath) 'exports'
            }
            New-Item -ItemType Directory -Force -Path $outDirEffective | Out-Null

            $revisions = New-Object System.Collections.ArrayList
            $rIndex = 0
            foreach ($rev in $doc.Revisions) {
                $rIndex++
                $author = ''; try { $author = [string]$rev.Author } catch {}
                $date = '';   try { $date = ([DateTime]$rev.Date).ToString('o') } catch {}
                $type = 0;    try { $type = [int]$rev.Type } catch {}
                $revText = ''; try { $revText = Safe-Text $rev.Range.Text } catch {}

                [void]$revisions.Add([ordered]@{
                    index  = $rIndex
                    author = $author
                    date   = $date
                    type   = $type
                    text   = $revText
                })
            }

            $redlinesPath = Join-Path $outDirEffective 'redlines.json'
            @{ sourcePath = $workingPath; revisionCount = $revisions.Count; revisions = $revisions } |
                ConvertTo-Json -Depth 6 |
                Set-Content -LiteralPath $redlinesPath -Encoding UTF8

            Write-Host ("Extracted {0} revisions to {1}" -f $revisions.Count, $redlinesPath)
        }
    }
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
