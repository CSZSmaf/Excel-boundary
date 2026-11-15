$ErrorActionPreference = 'Stop'
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) { throw 'Save and close Excel before uninstalling Excel Boundary.' }
$addinKey = 'HKCU:\Software\Microsoft\Office\Excel\Addins\Codex.ReadingSmooth'
$classKey = 'HKCU:\Software\Classes\CLSID\{6C31AE20-5B57-43E6-8984-790CFE5A5FE3}'
$progKey = 'HKCU:\Software\Classes\Codex.ReadingSmooth'
foreach ($key in $addinKey, $classKey, $progKey) {
    if (Test-Path -LiteralPath $key) { Remove-Item -LiteralPath $key -Recurse }
}
$installation = Join-Path $env:LOCALAPPDATA 'CodexExcelBoundary\ReadingAddin'
foreach ($name in 'ReadingSmooth.dll', 'ReadingBoundary.xll', 'settings.json') {
    $file = Join-Path $installation $name
    if (Test-Path -LiteralPath $file) { Remove-Item -LiteralPath $file }
}
Write-Output 'Excel Boundary has been unregistered and its installed binaries removed.'
