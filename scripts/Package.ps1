param([switch]$SkipBuild)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
if(-not $SkipBuild){& (Join-Path $PSScriptRoot 'build.ps1')}
$version='0.1.2'
$package=Join-Path $repo "dist\Excel-boundary-$version-windows-x64"
New-Item -ItemType Directory -Path $package,(Join-Path $package 'docs') -Force | Out-Null
foreach($name in 'ReadingSmooth.dll','ReadingBoundary.xll'){
    Copy-Item -LiteralPath (Join-Path $repo "build\release\$name") -Destination (Join-Path $package $name) -Force
}
foreach($name in 'Install.ps1','Uninstall.ps1','Register-Addin.ps1'){
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination (Join-Path $package $name) -Force
}
foreach($name in 'README.md','CHANGELOG.md','LICENSE','settings.example.json'){
    Copy-Item -LiteralPath (Join-Path $repo $name) -Destination (Join-Path $package $name) -Force
}
Copy-Item -Path (Join-Path $repo 'docs\*.md') -Destination (Join-Path $package 'docs') -Force
$archive=Join-Path $repo "dist\Excel-boundary-$version-windows-x64.zip"
Compress-Archive -Path (Join-Path $package '*') -DestinationPath $archive -Force
Write-Output $archive
