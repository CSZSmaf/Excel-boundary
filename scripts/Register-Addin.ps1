param([Parameter(Mandatory=$true)][string]$Folder, [int]$LoadBehavior = 3, [switch]$Lab)
$ErrorActionPreference = 'Stop'
$dll = Join-Path (Resolve-Path -LiteralPath $Folder).Path 'ReadingSmooth.dll'
$assemblyName = [Reflection.AssemblyName]::GetAssemblyName($dll).FullName
$classId = if ($Lab) { '{E510E056-1108-46AA-9F13-3A461F0AE244}' } else { '{6C31AE20-5B57-43E6-8984-790CFE5A5FE3}' }
$progId = if ($Lab) { 'ExcelBoundary.Lab' } else { 'Codex.ReadingSmooth' }
$classKey = 'HKCU:\Software\Classes\CLSID\' + $classId
$progKey = 'HKCU:\Software\Classes\' + $progId
$serverKey = "$classKey\InprocServer32"
New-Item -Path $serverKey -Force | Out-Null
Set-Item -LiteralPath $classKey -Value 'Excel Boundary'
Set-Item -LiteralPath $serverKey -Value 'mscoree.dll'
foreach ($entry in @{ThreadingModel='Both'; Class='ReadingAddin'; Assembly=$assemblyName; RuntimeVersion='v4.0.30319'; CodeBase=$dll}.GetEnumerator()) {
    New-ItemProperty -LiteralPath $serverKey -Name $entry.Key -Value $entry.Value -PropertyType String -Force | Out-Null
}
New-Item -Path "$classKey\ProgId", "$progKey\CLSID" -Force | Out-Null
Set-Item -LiteralPath "$classKey\ProgId" -Value $progId
Set-Item -LiteralPath "$progKey\CLSID" -Value $classId
$addinKey = 'HKCU:\Software\Microsoft\Office\Excel\Addins\' + $progId
New-Item -Path $addinKey -Force | Out-Null
New-ItemProperty -LiteralPath $addinKey -Name FriendlyName -Value 'Excel Boundary' -PropertyType String -Force | Out-Null
New-ItemProperty -LiteralPath $addinKey -Name Description -Value 'Native smooth scrolling with worksheet content boundaries.' -PropertyType String -Force | Out-Null
New-ItemProperty -LiteralPath $addinKey -Name LoadBehavior -Value $LoadBehavior -PropertyType DWord -Force | Out-Null
Write-Output "Registered $progId (LoadBehavior=$LoadBehavior)."
