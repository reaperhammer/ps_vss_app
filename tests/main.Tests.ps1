# Pester 5+ test suite for the ps_vss_app module.
# Run with: Invoke-Pester ./tests/main.Tests.ps1
# Or use the Build.ps1 wrapper which handles Pester installation and host selection.

BeforeAll {
    $moduleManifest = Join-Path $PSScriptRoot '..\ps_vss_app.psd1'
    if (-not (Test-Path $moduleManifest)) {
        throw "Module manifest not found at $moduleManifest. Run from the repo root."
    }
    Import-Module $moduleManifest -Force

    # Pester 5+ requires a command to exist before it can be mocked. Stub out
    # external commands referenced by the module so the mocks below have a
    # target. We declare these in the script scope so Pester's Mock can
    # intercept calls from inside the module.
    function global:vssadmin { }
    function global:diskshadow { }
    function global:cmd { }
    function global:'cmd.exe' { }
    function global:VssBackupHelper { }
}

AfterAll {
    Remove-Module ps_vss_app -ErrorAction SilentlyContinue
    Remove-Item function:global:vssadmin -ErrorAction SilentlyContinue
    Remove-Item function:global:diskshadow -ErrorAction SilentlyContinue
    Remove-Item function:global:cmd -ErrorAction SilentlyContinue
    Remove-Item function:global:'cmd.exe' -ErrorAction SilentlyContinue
    Remove-Item function:global:VssBackupHelper -ErrorAction SilentlyContinue
}

Describe 'ConvertTo-VSSCanonicalVolumeName' {
    It 'normalizes case and trailing slashes for volume device IDs' {
        InModuleScope ps_vss_app {
            $result = ConvertTo-VSSCanonicalVolumeName -Path '\\?\Volume{abc}\'
            $result | Should -Be '\\?\VOLUME{ABC}'
        }
    }

    It 'returns null for blank volume paths' {
        InModuleScope ps_vss_app {
            ConvertTo-VSSCanonicalVolumeName -Path '   ' | Should -BeNullOrEmpty
        }
    }

    It 'returns null for $null input' {
        InModuleScope ps_vss_app {
            ConvertTo-VSSCanonicalVolumeName -Path $null | Should -BeNullOrEmpty
        }
    }
}

Describe 'Get-VSSUsageColor' {
    It 'uses green below 70 percent' {
        Get-VSSUsageColor -UsagePercent 69.9 | Should -Be '#4CAF50'
    }

    It 'uses orange from 70 up to below 85 percent' {
        Get-VSSUsageColor -UsagePercent 70 | Should -Be '#FF9800'
        Get-VSSUsageColor -UsagePercent 84.9 | Should -Be '#FF9800'
    }

    It 'uses red from 85 percent upward' {
        Get-VSSUsageColor -UsagePercent 85 | Should -Be '#F44336'
        Get-VSSUsageColor -UsagePercent 100 | Should -Be '#F44336'
    }
}

Describe 'Get-VSSShadowStateName' {
    It 'maps numeric WMI state 12 to Created' {
        InModuleScope ps_vss_app {
            Get-VSSShadowStateName -State 12 | Should -Be 'Created'
        }
    }

    It 'keeps recognized textual state names' {
        InModuleScope ps_vss_app {
            Get-VSSShadowStateName -State 'Committed' | Should -Be 'Committed'
        }
    }

    It 'returns Unknown for null state' {
        InModuleScope ps_vss_app {
            Get-VSSShadowStateName -State $null | Should -Be 'Unknown'
        }
    }
}

Describe 'Get-VSSShadowStateColor' {
    It 'colors successful states green' {
        Get-VSSShadowStateColor -State 12 | Should -Be '#4CAF50'
        Get-VSSShadowStateColor -State 'Committed' | Should -Be '#4CAF50'
    }

    It 'colors failed states red' {
        Get-VSSShadowStateColor -State 13 | Should -Be '#F44336'
        Get-VSSShadowStateColor -State 'Failed' | Should -Be '#F44336'
    }
}

Describe 'Get-VSSAgeText' {
    It 'returns Unknown for null dates' {
        Get-VSSAgeText -DateTime $null | Should -Be 'Unknown'
    }

    It 'returns minute-level age text for recent dates' {
        Get-VSSAgeText -DateTime (Get-Date).AddMinutes(-5) | Should -Be '5m'
    }

    It 'returns hour-level age text for older dates' {
        Get-VSSAgeText -DateTime (Get-Date).AddHours(-3) | Should -Be '3h'
    }

    It 'returns day-level age text for very old dates' {
        Get-VSSAgeText -DateTime (Get-Date).AddDays(-2) | Should -Be '2d'
    }
}

Describe 'Get-VSSWriters' {
    It 'correctly parses mock vssadmin list writers output' {
        $vssadminOutput = @(
            "Writer name: 'Task Scheduler Writer'",
            "   Writer Id: {d61d61c8-d73a-4eee-8cdd-f6f9786b7124}",
            "   Writer Instance Id: {1b8dfd8e-128a-4c2f-b481-9b16ee892dcf}",
            "   State: [1] Stable",
            "   Last error: No error"
        )

        InModuleScope ps_vss_app -Parameters @{ Output = $vssadminOutput } {
            Mock -CommandName Invoke-NativeCommand -MockWith { param($CommandPath, $Arguments) return $Output }
            Mock -CommandName Test-Path -MockWith { return $true } -ParameterFilter { $Path -like '*vssadmin.exe' }
            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }

            $writers = @(Get-VSSWriters)

            $writers.Count | Should -Be 1
            $writers[0].WriterName | Should -Be 'Task Scheduler Writer'
            $writers[0].WriterId | Should -Be 'd61d61c8-d73a-4eee-8cdd-f6f9786b7124'
            $writers[0].State | Should -Be 'Stable'
            $writers[0].StateCode | Should -Be 1
            $writers[0].LastError | Should -Be 'No error'
            $writers[0].StatusColor | Should -Be '#4CAF50'
        }
    }

    It 'marks writers with errors red' {
        $failingOutput = @(
            "Writer name: 'Failing Writer'",
            "   Writer Id: {aaaa-bbbb-cccc}",
            "   State: [8] Failed",
            "   Last error: Timed out"
        )

        InModuleScope ps_vss_app -Parameters @{ Output = $failingOutput } {
            Mock -CommandName Invoke-NativeCommand -MockWith { param($CommandPath, $Arguments) return $Output }
            Mock -CommandName Test-Path -MockWith { return $true } -ParameterFilter { $Path -like '*vssadmin.exe' }
            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }

            $writers = @(Get-VSSWriters)
            $writers[0].StatusColor | Should -Be '#F44336'
            $writers[0].LastError | Should -Be 'Timed out'
        }
    }

    It 'falls back to COM-based enumeration when vssadmin parsing yields no writers' {
        $comOutput = @(
            '[{"name":"COM Writer 1","id":"{11111111-1111-1111-1111-111111111111}","instanceId":"{22222222-2222-2222-2222-222222222222}","state":"Unknown","stateCode":0,"instanceName":"","usage":""}]'
        )
        $unparseableOutput = @(
            "Schreibername: 'Irgendein Writer'",  # German "Writer name"
            "   Zustand: [1] Stabil"                  # German "State"
        )

        InModuleScope ps_vss_app -Parameters @{ ComOut = $comOutput; TextOut = $unparseableOutput } {
            Mock -CommandName Invoke-NativeCommand -MockWith {
                param($CommandPath, $Arguments)
                $global:LASTEXITCODE = 0
                if ($CommandPath -like '*vssadmin*') { return $TextOut }
                if ($Arguments -contains '--list-writers') { return $ComOut }
                return @()
            }
            Mock -CommandName Test-Path -MockWith { return $true } -ParameterFilter { $Path -like '*.exe' }
            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }

            $writers = @(Get-VSSWriters)
            $writers.Count | Should -Be 1
            $writers[0].WriterName | Should -Be 'COM Writer 1'
            $writers[0].WriterId | Should -Be '{11111111-1111-1111-1111-111111111111}'
        }
    }

    It 'skips vssadmin entirely when -ForceCom is specified' {
        $comOutput = @(
            '[{"name":"COM Writer","id":"{33333333-3333-3333-3333-333333333333}","instanceId":"{44444444-4444-4444-4444-444444444444}","state":"Unknown","stateCode":0,"instanceName":"","usage":""}]'
        )
        $vssadminCalled = $false

        InModuleScope ps_vss_app -Parameters @{ ComOut = $comOutput; Called = [ref]$vssadminCalled } {
            Mock -CommandName Invoke-NativeCommand -MockWith {
                param($CommandPath, $Arguments)
                $global:LASTEXITCODE = 0
                if ($CommandPath -like '*vssadmin*') { $Called.Value = $true; return @() }
                if ($Arguments -contains '--list-writers') { return $ComOut }
                return @()
            }
            Mock -CommandName Test-Path -MockWith { return $true } -ParameterFilter { $Path -like '*.exe' }
            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }

            $writers = @(Get-VSSWriters -ForceCom)
            $writers[0].WriterName | Should -Be 'COM Writer'
            $Called.Value | Should -BeFalse
        }
    }
}

Describe 'Get-VSSShadowStorage' {
    It 'correctly parses Win32_ShadowStorage WMI references' {
        InModuleScope ps_vss_app {
            $storageMock = [PSCustomObject]@{
                Volume = 'Win32_Volume.DeviceID="\\?\Volume{12345678-0000-0000-0000-100000000000}\"'
                DiffVolume = 'Win32_Volume.DeviceID="\\?\Volume{12345678-0000-0000-0000-100000000000}\"'
                MaxSpace = 10737418240
                AllocatedSpace = 5368709120
                UsedSpace = 1073741824
            }
            $volumeMock = [PSCustomObject]@{
                DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                DriveLetter = 'C:'
                VolumeName = 'Local Disk'
            }

            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }
            Mock -CommandName Get-VSSSupportedVolumes -MockWith { return @($volumeMock) }
            Mock -CommandName Get-WmiObject -MockWith {
                param($Class) if ($Class -eq 'Win32_ShadowStorage') { return @($storageMock) } else { return @() }
            }
            Mock -CommandName Resolve-VSSVolumePath -MockWith {
                [PSCustomObject]@{
                    DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                    CanonicalDevice = '\\?\VOLUME{12345678-0000-0000-0000-100000000000}'
                }
            }

            $storage = @(Get-VSSShadowStorage)
            $storage.Count | Should -Be 1
            $storage[0].VolumeLetter | Should -Be 'C:'
            $storage[0].DiffVolumeLetter | Should -Be 'C:'
            $storage[0].MaxSpaceGB | Should -Be 10
            $storage[0].AllocatedSpaceGB | Should -Be 5
            $storage[0].UsedSpaceGB | Should -Be 1
        }
    }

    It 'correctly parses Win32_ShadowStorage references with escaped backslashes' {
        InModuleScope ps_vss_app {
            $storageMock = [PSCustomObject]@{
                Volume = 'Win32_Volume.DeviceID="\\\\?\\Volume{12345678-0000-0000-0000-100000000000}\\"'
                DiffVolume = 'Win32_Volume.DeviceID="\\\\?\\Volume{12345678-0000-0000-0000-100000000000}\\"'
                MaxSpace = 10737418240
                AllocatedSpace = 5368709120
                UsedSpace = 1073741824
            }
            $volumeMock = [PSCustomObject]@{
                DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                DriveLetter = 'C:'
            }

            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }
            Mock -CommandName Get-VSSSupportedVolumes -MockWith { return @($volumeMock) }
            Mock -CommandName Get-WmiObject -MockWith {
                param($Class) if ($Class -eq 'Win32_ShadowStorage') { return @($storageMock) } else { return @() }
            }

            $storage = @(Get-VSSShadowStorage)
            $storage.Count | Should -Be 1
            $storage[0].VolumeLetter | Should -Be 'C:'
        }
    }

    It 'filters by VolumePath correctly even with escaped paths' {
        InModuleScope ps_vss_app {
            $storageMock = [PSCustomObject]@{
                Volume = 'Win32_Volume.DeviceID="\\\\?\\Volume{12345678-0000-0000-0000-100000000000}\\"'
                DiffVolume = 'Win32_Volume.DeviceID="\\\\?\\Volume{12345678-0000-0000-0000-100000000000}\\"'
                MaxSpace = 10737418240
                AllocatedSpace = 0
                UsedSpace = 0
            }
            $volumeMock = [PSCustomObject]@{
                DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                DriveLetter = 'C:'
                Label = 'Local Disk'
                FileSystem = 'NTFS'
                DriveType = 3
            }

            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }
            Mock -CommandName Get-VSSSupportedVolumes -MockWith { return @($volumeMock) }
            Mock -CommandName Get-WmiObject -MockWith {
                param($Class) if ($Class -eq 'Win32_ShadowStorage') { return @($storageMock) } else { return @() }
            }
            Mock -CommandName Resolve-VSSVolumePath -MockWith {
                [PSCustomObject]@{
                    DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                    CanonicalDevice = '\\?\VOLUME{12345678-0000-0000-0000-100000000000}'
                }
            }

            $storage = @(Get-VSSShadowStorage -VolumePath 'C:')
            $storage.Count | Should -Be 1
            $storage[0].VolumeLetter | Should -Be 'C:'
        }
    }
}

Describe 'Mount-VSSShadowCopy / Dismount-VSSShadowCopy' {
    It 'correctly mounts a shadow copy using mklink' {
        $cmdExecuted = $false

        InModuleScope ps_vss_app -Parameters @{ CmdExecFlag = [ref]$cmdExecuted } {
            $shadow = [PSCustomObject]@{
                ID = '{98765432-0000-0000-0000-100000000000}'
                DeviceObject = '\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1'
            }
            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }
            Mock -CommandName Get-WmiObject -MockWith { return @($shadow) }
            Mock -CommandName Invoke-NativeCommand -MockWith {
                param($CommandPath, $Arguments)
                $CmdExecFlag.Value = $true
                $global:LASTEXITCODE = 0
                return "symbolic link created"
            }

            $res = Mount-VSSShadowCopy -ShadowCopyID '{98765432-0000-0000-0000-100000000000}' -MountPath 'C:\mount'
            $res.Success | Should -BeTrue
            $res.MountPath | Should -Be 'C:\mount'
            $res.Target | Should -Be '\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1\'
            $CmdExecFlag.Value | Should -BeTrue
        }
    }

    It 'correctly dismounts a shadow copy using rmdir' {
        $cmdExecuted = $false
        InModuleScope ps_vss_app -Parameters @{ CmdExecFlag = [ref]$cmdExecuted } {
            Mock -CommandName Test-Path -MockWith { return $true }
            Mock -CommandName Invoke-NativeCommand -MockWith {
                param($CommandPath, $Arguments)
                $CmdExecFlag.Value = $true
                $global:LASTEXITCODE = 0
                return "directory removed"
            }

            $res = Dismount-VSSShadowCopy -MountPath 'C:\mount'
            $res.Success | Should -BeTrue
            $res.MountPath | Should -Be 'C:\mount'
            $CmdExecFlag.Value | Should -BeTrue
        }
    }
}

Describe 'New-VSSShadowCopy' {
    It 'handles Backup context failure with detailed error' {
        InModuleScope ps_vss_app {
            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }
            Mock -CommandName Get-VSSBackupHelperPath -MockWith { return $null }
            Mock -CommandName Test-Path -MockWith {
                param($Path) $false
            } -ParameterFilter { $Path -like '*diskshadow.exe' }
            Mock -CommandName Resolve-VSSVolumePath -MockWith {
                [PSCustomObject]@{
                    DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                    InputPath = 'C:'
                }
            }
            Mock -CommandName Get-WmiObject -MockWith {
                $mock = [PSCustomObject]@{}
                $mock | Add-Member -MemberType ScriptMethod -Name "Create" -Value {
                    param($volume, $context)
                    return [PSCustomObject]@{
                        ReturnValue = 5
                        ShadowID = $null
                    }
                }
                return $mock
            } -ParameterFilter { $List -eq $true }

            $res = New-VSSShadowCopy -VolumePath 'C:' -Context 'Backup' -Confirm:$false -ErrorAction SilentlyContinue
            $res.Success | Should -BeFalse
            $res.ReturnValue | Should -Be 5
            $res.ErrorDescription | Should -Match 'Unsupported shadow copy context'
        }
    }

    It 'uses VssBackupHelper when helper is available' {
        $helperOutput = @(
            'SUCCESS',
            'SnapshotID: {77777777-8888-9999-aaaa-bbbbbbbbbbbb}',
            'DeviceObject: \\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy7',
            'SnapshotSetID: {11111111-2222-3333-4444-555555555555}'
        )

        InModuleScope ps_vss_app -Parameters @{ HelperOutput = $helperOutput } {
            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }
            Mock -CommandName Get-VSSBackupHelperPath -MockWith { return 'VssBackupHelper' }
            Mock -CommandName Resolve-VSSVolumePath -MockWith {
                [PSCustomObject]@{
                    DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                    DriveLetter = 'C:'
                    InputPath = 'C:'
                }
            }
            Mock -CommandName Invoke-NativeCommand -MockWith {
                param($CommandPath, $Arguments)
                $global:LASTEXITCODE = 0
                return $HelperOutput
            }

            $res = New-VSSShadowCopy -VolumePath 'C:' -Context 'Backup' -Confirm:$false
            $res.Success | Should -BeTrue
            $res.ShadowID | Should -Be '{77777777-8888-9999-aaaa-bbbbbbbbbbbb}'
            $res.Context | Should -Be 'Backup'
        }
    }

    It 'piggybacks on diskshadow if helper is unavailable and Backup context is requested' {
        $diskShadowOutput = @(
            'Microsoft DiskShadow version 1.0',
            'Creating shadow copy...',
            '* Created shadow copy {885a0833-28eb-4e67-94a2-11c78479e0f6} on volume C:\'
        )
        $diskshadowExecuted = $false

        InModuleScope ps_vss_app -Parameters @{ Output = $diskShadowOutput; Flag = [ref]$diskshadowExecuted } {
            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }
            Mock -CommandName Get-VSSBackupHelperPath -MockWith { return $null }
            Mock -CommandName Resolve-VSSVolumePath -MockWith {
                [PSCustomObject]@{
                    DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                    DriveLetter = 'C:'
                    InputPath = 'C:'
                }
            }
            Mock -CommandName Test-Path -MockWith {
                param($Path) $Path -like '*diskshadow.exe'
            }
            Mock -CommandName Get-Command -MockWith {
                param($Name) if ($Name -eq 'diskshadow') { [PSCustomObject]@{ Name = 'diskshadow'; CommandType = 'ExternalScript' } } else { $null }
            }
            Mock -CommandName Invoke-NativeCommand -MockWith {
                param($CommandPath, $Arguments)
                if ($CommandPath -like '*diskshadow*' -or $CommandPath -eq 'diskshadow') {
                    $Flag.Value = $true
                    $global:LASTEXITCODE = 0
                    return $Output
                }
                return @()
            }

            $res = New-VSSShadowCopy -VolumePath 'C:' -Context 'Backup' -Confirm:$false
            $res.Success | Should -BeTrue
            $res.ShadowID | Should -Be '{885a0833-28eb-4e67-94a2-11c78479e0f6}'
            $Flag.Value | Should -BeTrue
        }
    }
}

Describe 'Set-VSSShadowStorageLimit' {
    It 'uses the temporary shadow copy workaround when direct creation fails' {
        $tempShadowCreated = $false
        $tempShadowDeleted = $false
        $limitApplied = $false
        $storageMockCount = 0

        InModuleScope ps_vss_app -Parameters @{
            TempCreated = [ref]$tempShadowCreated
            TempDeleted = [ref]$tempShadowDeleted
            LimitApplied = [ref]$limitApplied
            MockCount = [ref]$storageMockCount
        } {
            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }
            Mock -CommandName Resolve-VSSVolumePath -MockWith {
                [PSCustomObject]@{
                    DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                    CanonicalDevice = '\\?\VOLUME{12345678-0000-0000-0000-100000000000}'
                    DriveLetter = 'C:'
                    InputPath = 'C:'
                }
            }

            Mock -CommandName Get-WmiObject -MockWith {
                param($Class, $List)
                if ($Class -eq 'Win32_ShadowStorage') {
                    $MockCount.Value++
                    if ($MockCount.Value -eq 1) {
                        return @()
                    } else {
                        $mockStorage = [PSCustomObject]@{
                            Volume = 'Win32_Volume.DeviceID="\\\\?\\Volume{12345678-0000-0000-0000-100000000000}\\"'
                            DiffVolume = 'Win32_Volume.DeviceID="\\\\?\\Volume{12345678-0000-0000-0000-100000000000}\\"'
                            MaxSpace = 0
                        }
                        $mockStorage | Add-Member -MemberType ScriptMethod -Name "Put" -Value {
                            $LimitApplied.Value = $true
                            return [PSCustomObject]@{ Path = 'Win32_ShadowStorage' }
                        }
                        return @($mockStorage)
                    }
                } elseif ($Class -eq 'Win32_Volume') {
                    return @(
                        [PSCustomObject]@{
                            DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                            DriveLetter = 'C:'
                        }
                    )
                } else {
                    return $null
                }
            }

            # Force the vssadmin fallback to fail so the temporary shadow copy
            # workaround path is exercised.
            Mock -CommandName Invoke-NativeCommand -MockWith {
                param($CommandPath, $Arguments)
                $global:LASTEXITCODE = 1
                return @("Invalid command. vssadmin 1.0 - Volume Shadow Copy Service administrative command-line tool")
            }

            Mock -CommandName New-VSSShadowCopy -MockWith {
                $TempCreated.Value = $true
                [PSCustomObject]@{
                    Success = $true
                    ShadowID = '{98765432-0000-0000-0000-100000000000}'
                }
            }

            Mock -CommandName Remove-VSSShadowCopy -MockWith {
                $TempDeleted.Value = $true
                [PSCustomObject]@{ Success = $true }
            }

            $res = Set-VSSShadowStorageLimit -VolumePath 'C:' -MaxSpaceBytes 5368709120
            $res.Success | Should -BeTrue
            $TempCreated.Value | Should -BeTrue
            $LimitApplied.Value | Should -BeTrue
            $TempDeleted.Value | Should -BeTrue
        }
    }
}

Describe 'Export-VSSReport' {
    It 'exports volumes to CSV' {
        InModuleScope ps_vss_app {
            Mock -CommandName Get-VSSSupportedVolumes -MockWith {
                return @(
                    [PSCustomObject]@{ DriveLetter = 'C:'; CapacityGB = 100; FreeSpaceGB = 50 },
                    [PSCustomObject]@{ DriveLetter = 'D:'; CapacityGB = 200; FreeSpaceGB = 100 }
                )
            }
            $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ('vss-test-' + [guid]::NewGuid() + '.csv')
            try {
                $count = Export-VSSReport -Kind Volumes -Path $tempPath
                $count | Should -Be 2
                Test-Path $tempPath | Should -BeTrue
                $content = Get-Content $tempPath -Raw
                $content | Should -Match 'C:'
                $content | Should -Match 'D:'
            } finally {
                Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
            }
        }
    }

    It 'exports writers to JSON when -Format JSON' {
        InModuleScope ps_vss_app {
            Mock -CommandName Get-VSSWriters -MockWith {
                return @(
                    [PSCustomObject]@{ WriterName = 'W1'; State = 'Stable' }
                )
            }
            $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ('vss-test-' + [guid]::NewGuid() + '.json')
            try {
                $count = Export-VSSReport -Kind Writers -Path $tempPath -Format JSON
                $count | Should -Be 1
                $content = Get-Content $tempPath -Raw
                $content | Should -Match 'W1'
                $content | Should -Match 'Stable'
            } finally {
                Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
            }
        }
    }

    It 'throws when ShadowCopies kind is requested without VolumePath' {
        InModuleScope ps_vss_app {
            { Export-VSSReport -Kind ShadowCopies -Path 'C:\tmp.csv' } | Should -Throw '*VolumePath*'
        }
    }
}
