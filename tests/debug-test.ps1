Import-Module Pester -MinimumVersion 5.0 -Force
Import-Module 'F:\github\ps_vss_app_advancedvss\ps_vss_app\ps_vss_app.psd1' -Force

# Simulate the test setup
$comOutput = @(
    '[{"name":"COM Writer 1","id":"{11111111-1111-1111-1111-111111111111}","instanceId":"{22222222-2222-2222-2222-222222222222}","state":"Unknown","stateCode":0,"instanceName":"","usage":""}]'
)
$unparseableOutput = @(
    "Schreibername: 'Irgendein Writer'",
    "   Zustand: [1] Stabil"
)

& (Get-Module ps_vss_app) {
    Mock -CommandName Invoke-NativeCommand -MockWith {
        param($CommandPath, $Arguments)
        if ($CommandPath -like '*vssadmin*') { return $unparseableOutput }
        if ($Arguments -contains '--list-writers') { return $comOutput }
        return @()
    }
    Mock -CommandName Test-Path -MockWith { return $true } -ParameterFilter { $Path -like '*.exe' }
    Mock -CommandName Test-VSSAdministrator -MockWith { return $true }

    $writers = @(Get-VSSWriters)
    "Count: $($writers.Count)"
    foreach ($w in $writers) {
        "Name: $($w.WriterName) Id: $($w.WriterId)"
    }
} -Args $comOutput, $unparseableOutput
