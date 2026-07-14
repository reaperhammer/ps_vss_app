$ErrorActionPreference = 'Continue'
try {
    Install-Module -Name Pester -Force -SkipPublisherCheck -Scope CurrentUser -MinimumVersion 5.0 -Verbose -ErrorAction Continue 2>&1 | Out-Null
    'INSTALLED'
    Get-Module -ListAvailable -Name Pester | Sort-Object Version -Descending | Select-Object -First 3 | Format-Table Name, Version
} catch {
    'FAILED:'
    $_.Exception.Message
    $_.ScriptStackTrace
}
