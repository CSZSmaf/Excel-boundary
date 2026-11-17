param([Parameter(Mandatory=$true)][string]$Results)
$ErrorActionPreference='Stop'
$integration=Get-Content -LiteralPath (Join-Path $Results 'integration.json') -Raw | ConvertFrom-Json
$events=Get-Content -LiteralPath (Join-Path $Results 'boundary-native-events.jsonl') | ConvertFrom-Json | Where-Object {$_.pid -eq $integration.pid}
$body=@($events | Where-Object {$_.event -eq 'native-bound-probe' -and $_.sample -lt 600 -and $_.viewport[0] -gt 500 -and $_.viewport[1] -gt 400 -and $_.content[0] -lt 10000})
if($body.Count -ne 48){throw "Expected 48 native body samples, found $($body.Count)."}
foreach($sample in $body){
    for($axis=0;$axis -lt 2;$axis++){
        $viewportFactor=if($axis -eq 1){0.5}else{1.0}
        $expectedMax=[Math]::Max($sample.minimum[$axis],$sample.content[$axis]*$sample.scale-$sample.viewport[$axis]*$viewportFactor)
        if([Math]::Abs($sample.maximum[$axis]-$expectedMax) -gt 1){throw "Gesture hard boundary changed in sample $($sample.sample)."}
        if($sample.renderedOffset[$axis] -lt -$sample.maximum[$axis]-1 -or $sample.renderedOffset[$axis] -gt -$sample.minimum[$axis]+1){throw "Native boundary exceeded in sample $($sample.sample)."}
    }
}
$top=$body | Where-Object {$_.sample -eq 507}
if($top.minimum[1] -le 0 -or [Math]::Abs($top.renderedOffset[1]+$top.minimum[1]) -gt 1){throw 'Frozen top boundary validation failed.'}
$fractional=@($body | Where-Object {$_.sample -ge 100 -and $_.sample -lt 108 -and [Math]::Abs($_.renderedOffset[0]-[Math]::Round($_.renderedOffset[0])) -gt 0.02})
if($fractional.Count -eq 0){throw 'No native fractional scroll displacement observed.'}
$hit=Get-Content -LiteralPath (Join-Path $Results 'hit-test.json') -Raw | ConvertFrom-Json
if(-not $hit.passed){throw 'Rendered cell coordinates disagree with Excel hit testing.'}
$result=@{passed=$true;nativeSamples=$body.Count;frozenTop=$top.minimum[1];fractionalMotionSamples=$fractional.Count;hitTestCases=$hit.results.Count;inactiveTimerCallbacks=$integration.inactiveTimerCallbacks;entryVerticalFraction=$integration.entryVerticalFraction}
$result | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $Results 'native-validation.json') -Encoding UTF8
$result | ConvertTo-Json
