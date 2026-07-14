Import-Module 'F:\github\ps_vss_app_advancedvss\ps_vss_app\ps_vss_app.psd1' -Force
foreach ($cmd in Get-Command -Module ps_vss_app) {
    $help = Get-Help $cmd.Name -ErrorAction SilentlyContinue
    $hasDesc = $false
    if ($help -and $help.description) {
        if ($help.description.Text -and $help.description.Text.Trim()) {
            $hasDesc = $true
        } elseif ($help.description -is [string] -and $help.description.Trim()) {
            $hasDesc = $true
        }
    }
    $status = if ($hasDesc) { "OK" } else { "MISSING" }
    "{0,-30} {1}" -f $cmd.Name, $status
}
