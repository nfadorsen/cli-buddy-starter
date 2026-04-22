[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$Path
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $Path)) {
    Write-Error "File not found: $Path"
    exit 2
}

$resolved = (Resolve-Path -LiteralPath $Path).Path

$bytes = New-Object byte[] 8
$fs = [System.IO.File]::Open($resolved, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
try {
    $read = $fs.Read($bytes, 0, 8)
} finally {
    $fs.Close()
    $fs.Dispose()
}

$headerHex = ($bytes[0..([Math]::Min($read-1,7))] | ForEach-Object { $_.ToString('X2') }) -join ' '

$containerType = 'unknown'
if ($read -ge 2 -and $bytes[0] -eq 0x50 -and $bytes[1] -eq 0x4B) {
    $containerType = 'zip'
} elseif ($read -ge 4 -and $bytes[0] -eq 0xD0 -and $bytes[1] -eq 0xCF -and $bytes[2] -eq 0x11 -and $bytes[3] -eq 0xE0) {
    $containerType = 'ole2'
}

$obj = [ordered]@{
    path          = $resolved
    headerHex     = $headerHex
    containerType = $containerType
}

$obj | ConvertTo-Json -Compress
