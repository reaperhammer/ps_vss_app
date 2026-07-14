Import-Module Pester -MinimumVersion 5.0 -Force

# Run only the Get-VSSWriters tests
$config = New-PesterConfiguration
$config.Run.Path = "F:\github\ps_vss_app_advancedvss\ps_vss_app\tests\main.Tests.ps1"
$config.Filter.FullName = '*Get-VSSWriters*falls back*'
$config.Output.Verbosity = 'Detailed'
$config.Run.PassThru = $true
$config.Run.Exit = $false

$result = Invoke-Pester -Configuration $config
"---"
"Passed: $($result.PassedCount)  Failed: $($result.FailedCount)"
"---"
$result.Failed | ForEach-Object {
    "FAILED: $($_.ExpandedPath)"
    "Error: $($_.ErrorRecord.Exception.Message)"
    "Stack: $($_.ErrorRecord.Exception.StackTrace)"
}
