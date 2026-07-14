# Build.ps1 - Test runner for the ps_vss_app module.
#
# Installs Pester 5+ if needed and runs the test suite. Designed to be run from
# the repo root in Windows PowerShell 5.1:
#
#     .\Build.ps1
#
# On a CI runner (Linux/macOS/Windows with PowerShell 7+) Pester 5 is usually
# pre-installed; on a stock Windows PowerShell 5.1 host the script installs it
# into the current user's scope so no admin rights are required.

[CmdletBinding()]
param(
    [string]$PesterVersion = '5.7.0',
    [switch]$CI
)

$ErrorActionPreference = 'Stop'

# Pester 5+ requires PowerShell 5.1 or later. The GUI/CLI is built for 5.1, so we
# default to that; passing -CI uses whatever PowerShell is on the path (which is
# useful in GitHub Actions where the runner is PowerShell 7).
$psExe = if ($CI) { (Get-Process -Id $PID).Path } else {
    Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
}
if (-not $CI -and -not (Test-Path $psExe)) {
    throw "Windows PowerShell 5.1 was not found at $psExe."
}

Write-Host "Using PowerShell: $psExe" -ForegroundColor Cyan

$psArgs = @(
    '-NoProfile'
    '-ExecutionPolicy', 'Bypass'
    '-File', $PSCommandPath
    '-PesterVersion', $PesterVersion
)

# Detect whether we're already inside a 5.1 host; if so, run the tests inline so
# the user gets visible output. Otherwise relaunch.
$needsRelaunch = -not $CI -and $PSVersionTable.PSVersion.Major -lt 5
if ($needsRelaunch) {
    Write-Host "Relaunching in Windows PowerShell 5.1..." -ForegroundColor Yellow
    & $psExe @psArgs
    exit $LASTEXITCODE
}

Write-Host "Pester version: $($PSVersionTable.PSVersion)" -ForegroundColor Cyan

# Install Pester 5+ if not present, or if an older version is installed.
$installed = Get-Module -ListAvailable -Name Pester | Sort-Object Version -Descending | Select-Object -First 1
if (-not $installed -or $installed.Version.Major -lt 5) {
    Write-Host "Installing Pester $PesterVersion..." -ForegroundColor Yellow
    Install-Module -Name Pester -Force -SkipPublisherCheck -Scope CurrentUser -MinimumVersion $PesterVersion
    Remove-Module Pester -ErrorAction SilentlyContinue
    Import-Module Pester -MinimumVersion $PesterVersion -Force
} else {
    Write-Host "Pester $($installed.Version) already installed." -ForegroundColor Green
}

# Resolve repo root (parent of this script).
$repoRoot = Split-Path -Parent $PSCommandPath
$testsPath = Join-Path $repoRoot 'tests/main.Tests.ps1'

Write-Host "Running tests in $testsPath..." -ForegroundColor Cyan
$config = New-PesterConfiguration
$config.Run.Path = $testsPath
$config.Output.Verbosity = 'Detailed'
$config.TestResult.Enabled = $true
$config.TestResult.OutputPath = Join-Path $repoRoot 'tests/TestResults.xml'
$config.Run.Exit = if ($CI) { $true } else { $false }

$result = Invoke-Pester -Configuration $config
if ($result.FailedCount -gt 0) {
    Write-Host "`n$($result.FailedCount) test(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host "`nAll $($result.PassedCount) test(s) passed." -ForegroundColor Green
exit 0
