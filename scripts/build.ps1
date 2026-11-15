param([switch]$Lab, [string]$ExcelInteropPath)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$configuration = if ($Lab) { 'lab' } else { 'release' }
$output = Join-Path $repo "build\$configuration"
New-Item -ItemType Directory -Path $output -Force | Out-Null

$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$visualStudio = & $vswhere -latest -prerelease -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
if (-not $visualStudio) { throw 'Install Visual Studio Build Tools with Desktop development with C++.' }
$vcVersion = (Get-Content -LiteralPath (Join-Path $visualStudio 'VC\Auxiliary\Build\Microsoft.VCToolsVersion.default.txt') -Raw).Trim()
$vc = Join-Path $visualStudio "VC\Tools\MSVC\$vcVersion"
$kits = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows Kits\Installed Roots').KitsRoot10
$sdkVersion = Get-ChildItem -LiteralPath (Join-Path $kits 'Include') -Directory | Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'cppwinrt\winrt\Windows.UI.Composition.h') } | Sort-Object { [version]$_.Name } -Descending | Select-Object -First 1 -ExpandProperty Name
if (-not $sdkVersion) { throw 'Install the Windows SDK with C++/WinRT headers.' }
$compiler = Join-Path $vc 'bin\Hostx64\x64\cl.exe'
$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'

if (-not $ExcelInteropPath) {
    $gac = Join-Path $env:WINDIR 'assembly\GAC_MSIL\Microsoft.Office.Interop.Excel'
    if (Test-Path -LiteralPath $gac) { $ExcelInteropPath = Get-ChildItem -LiteralPath $gac -Filter Microsoft.Office.Interop.Excel.dll -Recurse | Select-Object -First 1 -ExpandProperty FullName }
}
if (-not $ExcelInteropPath) {
    $packageDir = Join-Path $repo 'packages\excel-interop'
    $packageZip = Join-Path $repo 'packages\excel-interop.zip'
    New-Item -ItemType Directory -Path $packageDir -Force | Out-Null
    if (-not (Test-Path -LiteralPath (Join-Path $packageDir 'lib\net20\Microsoft.Office.Interop.Excel.dll'))) {
        Invoke-WebRequest -Uri 'https://api.nuget.org/v3-flatcontainer/microsoft.office.interop.excel/15.0.4795.1001/microsoft.office.interop.excel.15.0.4795.1001.nupkg' -OutFile $packageZip
        Expand-Archive -LiteralPath $packageZip -DestinationPath $packageDir -Force
    }
    $ExcelInteropPath = Get-ChildItem -LiteralPath $packageDir -Filter Microsoft.Office.Interop.Excel.dll -Recurse | Select-Object -First 1 -ExpandProperty FullName
}
if (-not (Test-Path -LiteralPath $ExcelInteropPath)) { throw 'Excel interop assembly was not found.' }
$oldInclude, $oldLib = $env:INCLUDE, $env:LIB
try {
    $env:INCLUDE = @("$vc\include", "$kits\Include\$sdkVersion\cppwinrt", "$kits\Include\$sdkVersion\ucrt", "$kits\Include\$sdkVersion\shared", "$kits\Include\$sdkVersion\um", "$kits\Include\$sdkVersion\winrt") -join ';'
    $env:LIB = @("$vc\lib\x64", "$kits\Lib\$sdkVersion\ucrt\x64", "$kits\Lib\$sdkVersion\um\x64") -join ';'
    $nativeFlags = @('/nologo', '/std:c++17', '/utf-8', '/EHsc', '/O2', '/MT', '/guard:cf', '/DUNICODE', '/D_UNICODE', '/DCODEX_NATIVE_ADDIN')
    if ($Lab) { $nativeFlags += @('/DCODEX_NATIVE_LAB', '/DCODEX_NATIVE_PROFILE') }
    & $compiler @nativeFlags /LD (Join-Path $repo 'src\native\BoundaryExpressions.cpp') "/Fo$output\BoundaryExpressions.obj" /link "/OUT:$output\ReadingBoundary.xll" "/IMPLIB:$output\ReadingBoundary.lib" /guard:cf user32.lib ole32.lib runtimeobject.lib windowsapp.lib CoreMessaging.lib
    if ($LASTEXITCODE) { throw 'Native build failed.' }
    $managedFlags = @('/nologo', '/target:library', '/platform:x64', '/optimize+', '/reference:Microsoft.CSharp.dll', '/reference:System.Windows.Forms.dll', '/reference:System.Web.Extensions.dll', "/link:$ExcelInteropPath", "/out:$output\ReadingSmooth.dll")
    if ($Lab) { $managedFlags += '/define:BOUNDARY_LAB' }
    $sources = Get-ChildItem -LiteralPath (Join-Path $repo 'src\managed') -Filter '*.cs' | Select-Object -ExpandProperty FullName
    & $csc @managedFlags @sources
    if ($LASTEXITCODE) { throw 'Managed build failed.' }
    & $compiler /nologo /std:c++17 /utf-8 /EHsc /O2 /MT /DCODEX_NATIVE_ADDIN (Join-Path $repo 'tests\NativeExpressions.cpp') "/Fo$output\NativeExpressions.obj" /link "/OUT:$output\NativeExpressions.exe"
    if ($LASTEXITCODE) { throw 'Native test build failed.' }
    & $csc /nologo /target:exe /platform:x64 "/out:$output\ScheduleTests.exe" (Join-Path $repo 'tests\ScheduleTests.cs') (Join-Path $repo 'src\managed\MaintenanceSchedule.cs')
    if ($LASTEXITCODE) { throw 'Scheduler test build failed.' }
    & (Join-Path $output 'NativeExpressions.exe')
    if ($LASTEXITCODE) { throw 'Native expression tests failed.' }
    & (Join-Path $output 'ScheduleTests.exe')
    if ($LASTEXITCODE) { throw 'Scheduler tests failed.' }
    if ($Lab) {
        & $compiler @nativeFlags (Join-Path $repo 'tests\PrivateDesktop.cpp') "/Fo$output\PrivateDesktop.obj" /link "/OUT:$output\PrivateDesktop.exe" user32.lib
        if ($LASTEXITCODE) { throw 'Private desktop runner build failed.' }
        & $csc /nologo /target:winexe /platform:x64 /debug+ /reference:Microsoft.CSharp.dll /reference:System.Windows.Forms.dll /reference:System.Web.Extensions.dll "/out:$output\ExcelIntegration.exe" (Join-Path $repo 'tests\ExcelIntegration.cs')
        if ($LASTEXITCODE) { throw 'Excel integration test build failed.' }
    }
} finally { $env:INCLUDE = $oldInclude; $env:LIB = $oldLib }
Write-Output "Built $output"
