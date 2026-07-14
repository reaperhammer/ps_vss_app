Import-Module Pester -MinimumVersion 5.0 -Force
$config = New-PesterConfiguration
$config.Run.Path = $PSScriptRoot + '/main.Tests.ps1'
$config.Output.Verbosity = 'Detailed'
$config.Run.PassThru = $true
$config.Run.Exit = $false

$result = Invoke-Pester -Configuration $config
"Passed: $($result.PassedCount)  Failed: $($result.FailedCount)  Skipped: $($result.SkippedCount)  Total: $($result.TotalCount)"

if ($result.FailedCount -gt 0) { exit 1 } else { exit 0 }
