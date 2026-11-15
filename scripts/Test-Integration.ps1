param([switch]$SkipBuild)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
if(-not $SkipBuild){& (Join-Path $PSScriptRoot 'build.ps1') -Lab}
$lab=Join-Path $repo 'build\lab'
$results=Join-Path $repo 'test-results'
New-Item -ItemType Directory -Path $results -Force | Out-Null
$env:EXCEL_BOUNDARY_LOG_DIR=$results
$production='HKCU:\Software\Microsoft\Office\Excel\Addins\Codex.ReadingSmooth'
$previous=$null
if(Test-Path -LiteralPath $production){$previous=Get-ItemPropertyValue -LiteralPath $production -Name LoadBehavior}
try {
    & (Join-Path $PSScriptRoot 'Register-Addin.ps1') -Folder $lab -Lab -LoadBehavior 3
    if($null -ne $previous){$env:EXCEL_BOUNDARY_PREVIOUS_LOAD=[string]$previous;Set-ItemProperty -LiteralPath $production -Name LoadBehavior -Value 0}
    & (Join-Path $lab 'PrivateDesktop.exe')
    if($LASTEXITCODE){throw "Excel integration failed. See $results\error.txt"}
    Get-Content -LiteralPath (Join-Path $results 'integration.json')
    & (Join-Path $PSScriptRoot 'Validate-Native.ps1') -Results $results
} finally {
    if($null -ne $previous){Set-ItemProperty -LiteralPath $production -Name LoadBehavior -Value $previous}
    Remove-Item Env:\EXCEL_BOUNDARY_PREVIOUS_LOAD -ErrorAction SilentlyContinue
    Remove-Item Env:\EXCEL_BOUNDARY_LOG_DIR -ErrorAction SilentlyContinue
    foreach($key in 'HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelBoundary.Lab','HKCU:\Software\Classes\ExcelBoundary.Lab','HKCU:\Software\Classes\CLSID\{E510E056-1108-46AA-9F13-3A461F0AE244}'){
        if(Test-Path -LiteralPath $key){Remove-Item -LiteralPath $key -Recurse}
    }
}
