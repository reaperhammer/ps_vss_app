$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $here
$sut = Join-Path $repoRoot 'main.ps1'
. $sut

Describe 'VSS helper functions' {
    Context 'Volume path normalization' {
        It 'normalizes case and trailing slashes for volume device IDs' {
            ConvertTo-VSSCanonicalVolumeName -Path '\\?\Volume{abc}\' | Should Be '\\?\VOLUME{ABC}'
        }

        It 'returns null for blank volume paths' {
            ConvertTo-VSSCanonicalVolumeName -Path '   ' | Should BeNullOrEmpty
        }
    }

    Context 'Volume usage display' {
        It 'uses green below 70 percent' {
            Get-VSSUsageColor -UsagePercent 69.9 | Should Be '#4CAF50'
        }

        It 'uses orange from 70 up to below 85 percent' {
            Get-VSSUsageColor -UsagePercent 70 | Should Be '#FF9800'
            Get-VSSUsageColor -UsagePercent 84.9 | Should Be '#FF9800'
        }

        It 'uses red from 85 percent upward' {
            Get-VSSUsageColor -UsagePercent 85 | Should Be '#F44336'
        }
    }

    Context 'Shadow copy state display' {
        It 'maps numeric WMI state 12 to Created' {
            Get-VSSShadowStateName -State 12 | Should Be 'Created'
        }

        It 'keeps recognized textual state names' {
            Get-VSSShadowStateName -State 'Committed' | Should Be 'Committed'
        }

        It 'colors successful states green' {
            Get-VSSShadowStateColor -State 12 | Should Be '#4CAF50'
            Get-VSSShadowStateColor -State 'Committed' | Should Be '#4CAF50'
        }

        It 'colors failed states red' {
            Get-VSSShadowStateColor -State 13 | Should Be '#F44336'
            Get-VSSShadowStateColor -State 'Failed' | Should Be '#F44336'
        }
    }

    Context 'Shadow copy age display' {
        It 'returns Unknown for null dates' {
            Get-VSSAgeText -DateTime $null | Should Be 'Unknown'
        }

        It 'returns minute-level age text for recent dates' {
            Get-VSSAgeText -DateTime (Get-Date).AddMinutes(-5) | Should Be '5m'
        }
    }
}
