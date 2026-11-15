param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [string]$WorksheetName = '',
    [string[]]$IgnoredHeaders = @(),
    [switch]$Diagnostics,
    [string]$BinaryDirectory = $PSScriptRoot
)
$ErrorActionPreference = 'Stop'
if (-not [Environment]::Is64BitProcess) { throw 'Run this installer from 64-bit PowerShell.' }
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) { throw 'Save and close Excel before installing or updating Excel Boundary.' }
$excelPath = $null
foreach ($key in 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe') {
    if (Test-Path -LiteralPath $key) { $excelPath = (Get-Item -LiteralPath $key).GetValue(''); break }
}
if (-not $excelPath) { throw 'Desktop Excel was not found.' }
$version = (Get-Item -LiteralPath $excelPath).VersionInfo.FileVersion
if ($version -ne '16.0.20527.20196') { throw "This release supports Excel 16.0.20527.20196. Installed: $version" }
$reader = New-Object IO.BinaryReader([IO.File]::OpenRead($excelPath))
try {
    $reader.BaseStream.Position=0x3c
    $pe=$reader.ReadInt32()
    $reader.BaseStream.Position=$pe+4
    if ($reader.ReadUInt16() -ne 0x8664) { throw 'This release requires 64-bit Excel.' }
} finally { $reader.Dispose() }
$workbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
if (-not (Test-Path -LiteralPath (Join-Path $BinaryDirectory 'ReadingSmooth.dll'))) {
    $BinaryDirectory = Join-Path (Split-Path $PSScriptRoot -Parent) 'build\release'
}
$installation = Join-Path $env:LOCALAPPDATA 'CodexExcelBoundary\ReadingAddin'
New-Item -ItemType Directory -Path $installation -Force | Out-Null
foreach ($name in 'ReadingSmooth.dll', 'ReadingBoundary.xll') {
    Copy-Item -LiteralPath (Join-Path $BinaryDirectory $name) -Destination (Join-Path $installation $name) -Force
    Unblock-File -LiteralPath (Join-Path $installation $name)
}
@{WorkbookPath=$workbook; WorksheetName=$WorksheetName; IgnoredHeaders=@($IgnoredHeaders); Diagnostics=[bool]$Diagnostics} |
    ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $installation 'settings.json') -Encoding UTF8
& (Join-Path $PSScriptRoot 'Register-Addin.ps1') -Folder $installation
Write-Output "Installed. Open your workbook normally: $workbook"
