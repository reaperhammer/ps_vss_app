$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $here
$sut = Join-Path $repoRoot 'main.ps1'

Describe 'VSS helper functions' {
    # Define a placeholder function for vssadmin so Pester 3.x can mock it successfully
    function vssadmin {}
    function diskshadow {}
    . $sut



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

    Context 'VSS Writers parsing' {
        It 'correctly parses mock vssadmin list writers output' {
            Mock -CommandName vssadmin -MockWith {
                return @(
                    "Writer name: 'Task Scheduler Writer'",
                    "   Writer Id: {d61d61c8-d73a-4eee-8cdd-f6f9786b7124}",
                    "   Writer Instance Id: {1b8dfd8e-128a-4c2f-b481-9b16ee892dcf}",
                    "   State: [1] Stable",
                    "   Last error: No error"
                )
            }
            Mock -CommandName Test-Path -MockWith { return $true }
            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }

            $writers = @(Get-VSSWriters)
            $writers.Count | Should Be 1
            $writers[0].WriterName | Should Be 'Task Scheduler Writer'
            $writers[0].WriterId | Should Be 'd61d61c8-d73a-4eee-8cdd-f6f9786b7124'
            $writers[0].State | Should Be 'Stable'
            $writers[0].StateCode | Should Be 1
            $writers[0].LastError | Should Be 'No error'
            $writers[0].StatusColor | Should Be '#4CAF50'
        }
    }

    Context 'VSS Shadow Storage query' {
        It 'correctly parses Win32_ShadowStorage WMI references' {
            Mock -CommandName Get-WmiObject -MockWith {
                if ($Class -eq 'Win32_ShadowStorage') {
                    return @(
                        [PSCustomObject]@{
                            Volume = 'Win32_Volume.DeviceID="\\?\Volume{12345678-0000-0000-0000-100000000000}\"'
                            DiffVolume = 'Win32_Volume.DeviceID="\\?\Volume{12345678-0000-0000-0000-100000000000}\"'
                            MaxSpace = 10737418240 # 10GB
                            AllocatedSpace = 5368709120 # 5GB
                            UsedSpace = 1073741824 # 1GB
                        }
                    )
                }
            }
            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }
            Mock -CommandName Get-VSSSupportedVolumes -MockWith {
                return @(
                    [PSCustomObject]@{
                        DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                        DriveLetter = 'C:'
                        VolumeName = 'Local Disk'
                    }
                )
            }

            $storage = @(Get-VSSShadowStorage)
            $storage.Count | Should Be 1
            $storage[0].VolumeLetter | Should Be 'C:'
            $storage[0].DiffVolumeLetter | Should Be 'C:'
            $storage[0].MaxSpaceGB | Should Be 10
            $storage[0].AllocatedSpaceGB | Should Be 5
            $storage[0].UsedSpaceGB | Should Be 1
        }

        It 'correctly parses Win32_ShadowStorage references with escaped backslashes' {
            Mock -CommandName Get-WmiObject -MockWith {
                if ($Class -eq 'Win32_ShadowStorage') {
                    return @(
                        [PSCustomObject]@{
                            Volume = 'Win32_Volume.DeviceID="\\\\?\\Volume{12345678-0000-0000-0000-100000000000}\\"'
                            DiffVolume = 'Win32_Volume.DeviceID="\\\\?\\Volume{12345678-0000-0000-0000-100000000000}\\"'
                            MaxSpace = 10737418240 # 10GB
                            AllocatedSpace = 5368709120 # 5GB
                            UsedSpace = 1073741824 # 1GB
                        }
                    )
                }
            }
            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }
            Mock -CommandName Get-VSSSupportedVolumes -MockWith {
                return @(
                    [PSCustomObject]@{
                        DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                        DriveLetter = 'C:'
                        VolumeName = 'Local Disk'
                    }
                )
            }

            $storage = @(Get-VSSShadowStorage)
            $storage.Count | Should Be 1
            $storage[0].VolumeLetter | Should Be 'C:'
            $storage[0].DiffVolumeLetter | Should Be 'C:'
            $storage[0].MaxSpaceGB | Should Be 10
        }

        It 'filters by VolumePath correctly even with escaped paths' {
            Mock -CommandName Get-WmiObject -MockWith {
                if ($Class -eq 'Win32_ShadowStorage') {
                    return @(
                        [PSCustomObject]@{
                            Volume = 'Win32_Volume.DeviceID="\\\\?\\Volume{12345678-0000-0000-0000-100000000000}\\"'
                            DiffVolume = 'Win32_Volume.DeviceID="\\\\?\\Volume{12345678-0000-0000-0000-100000000000}\\"'
                            MaxSpace = 10737418240
                            AllocatedSpace = 0
                            UsedSpace = 0
                        }
                    )
                }
                elseif ($Class -eq 'Win32_Volume') {
                    return @(
                        [PSCustomObject]@{
                            DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                            DriveLetter = 'C:'
                            Label = 'Local Disk'
                            FileSystem = 'NTFS'
                            DriveType = 3
                        }
                    )
                }
            }
            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }
            Mock -CommandName Get-VSSSupportedVolumes -MockWith {
                return @(
                    [PSCustomObject]@{
                        DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                        DriveLetter = 'C:'
                        VolumeName = 'Local Disk'
                    }
                )
            }

            $storage = @(Get-VSSShadowStorage -VolumePath 'C:')
            $storage.Count | Should Be 1
            $storage[0].VolumeLetter | Should Be 'C:'
        }
    }

    Context 'VSS Shadow Copy Mount & Dismount' {
        It 'correctly mounts a shadow copy using mklink' {
            Mock -CommandName Get-WmiObject -MockWith {
                return @(
                    [PSCustomObject]@{
                        ID = '{98765432-0000-0000-0000-100000000000}'
                        DeviceObject = '\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1'
                    }
                )
            }
            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }
            Mock -CommandName cmd.exe -MockWith {
                $script:cmdExecuted = $true
                $global:LASTEXITCODE = 0
                return "symbolic link created"
            }
            Mock -CommandName Test-Path -MockWith { if ($Path -eq 'C:\' -or $Path -eq 'C:') { return $true } else { return $false } }

            $res = Mount-VSSShadowCopy -ShadowCopyID '{98765432-0000-0000-0000-100000000000}' -MountPath 'C:\mount'
            $res.Success | Should Be $true
            $res.MountPath | Should Be 'C:\mount'
            $res.Target | Should Be '\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1\'
            $script:cmdExecuted | Should Be $true
        }

        It 'correctly dismounts a shadow copy using rmdir' {
            Mock -CommandName cmd.exe -MockWith {
                $script:cmdExecuted = $true
                $global:LASTEXITCODE = 0
                return "directory removed"
            }
            Mock -CommandName Test-Path -MockWith { return $true }

            $res = Dismount-VSSShadowCopy -MountPath 'C:\mount'
            $res.Success | Should Be $true
            $res.MountPath | Should Be 'C:\mount'
            $script:cmdExecuted | Should Be $true
        }
    }

    Context 'VSS Shadow Copy Creation' {
        It 'handles Backup context failure with detailed error' {
            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }
            Mock -CommandName Get-VSSBackupHelperPath -MockWith { return $null }
            Mock -CommandName Resolve-VSSVolumePath -MockWith {
                return [PSCustomObject]@{
                    DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                }
            }
            Mock -CommandName Get-WmiObject -MockWith {
                if ($Class -eq 'Win32_ShadowCopy') {
                    $mock = [PSCustomObject]@{}
                    $mock | Add-Member -MemberType ScriptMethod -Name "Create" -Value {
                        param($volume, $context)
                        return [PSCustomObject]@{
                            ReturnValue = 5
                            ShadowID = $null
                        }
                    }
                    return $mock
                }
            }

            # New-VSSShadowCopy will write an error, we can check that it returns the object with ErrorDescription
            $res = New-VSSShadowCopy -VolumePath 'C:' -Context 'Backup' -Confirm:$false
            $res.Success | Should Be $false
            $res.ReturnValue | Should Be 5
            $res.ErrorDescription | Should Match "Unsupported shadow copy context"
        }

        It 'piggybacks on diskshadow if available and Backup context is requested' {
            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }
            Mock -CommandName Resolve-VSSVolumePath -MockWith {
                return [PSCustomObject]@{
                    DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                    DriveLetter = 'C:'
                }
            }
            
            Mock -CommandName Test-Path -MockWith {
                if ($Path -like "*diskshadow.exe") { return $true }
                return $false
            }
            
            Mock -CommandName Get-Command -MockWith {
                if ($Name -eq 'diskshadow') {
                    return [PSCustomObject]@{ Name = 'diskshadow'; CommandType = 'ExternalScript' }
                }
                return $null
            }
            
            $script:diskshadowExecuted = $false
            Mock -CommandName diskshadow -MockWith {
                $script:diskshadowExecuted = $true
                $global:LASTEXITCODE = 0
                return @(
                    "Microsoft DiskShadow version 1.0",
                    "Creating shadow copy...",
                    "* Created shadow copy {885a0833-28eb-4e67-94a2-11c78479e0f6} on volume C:\"
                )
            }

            $res = New-VSSShadowCopy -VolumePath 'C:' -Context 'Backup' -Confirm:$false
            $res.Success | Should Be $true
            $res.ShadowID | Should Be '{885a0833-28eb-4e67-94a2-11c78479e0f6}'
            $script:diskshadowExecuted | Should Be $true
        }

        It 'uses VssBackupHelper if helper is available' {
            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }
            Mock -CommandName Resolve-VSSVolumePath -MockWith {
                return [PSCustomObject]@{
                    DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                    DriveLetter = 'C:'
                }
            }
            
            function MockHelper {
                $global:LASTEXITCODE = 0
                return @(
                    "SUCCESS",
                    "SnapshotID: {77777777-8888-9999-aaaa-bbbbbbbbbbbb}",
                    "DeviceObject: \\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy7",
                    "SnapshotSetID: {11111111-2222-3333-4444-555555555555}"
                )
            }
            
            Mock -CommandName Get-VSSBackupHelperPath -MockWith {
                return "MockHelper"
            }

            $res = New-VSSShadowCopy -VolumePath 'C:' -Context 'Backup' -Confirm:$false
            $res.Success | Should Be $true
            $res.ShadowID | Should Be '{77777777-8888-9999-aaaa-bbbbbbbbbbbb}'
            $res.Context | Should Be 'Backup'
        }
    }

    Context 'VSS Shadow Storage Limit configuration' {
        It 'uses the temporary shadow copy workaround when direct creation fails' {
            Mock -CommandName Test-VSSAdministrator -MockWith { return $true }
            Mock -CommandName Resolve-VSSVolumePath -MockWith {
                return [PSCustomObject]@{
                    DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                    CanonicalDevice = '\\?\VOLUME{12345678-0000-0000-0000-100000000000}'
                    DriveLetter = 'C:'
                }
            }
            
            $script:tempShadowCreated = $false
            $script:tempShadowDeleted = $false
            $script:limitApplied = $false

            # Mock Get-WmiObject to return no storage at first (so it attempts to create)
            # but return the storage when searched inside the workaround
            $script:storageMockCount = 0
            Mock -CommandName Get-WmiObject -MockWith {
                if ($Class -eq 'Win32_ShadowStorage') {
                    $script:storageMockCount++
                    if ($script:storageMockCount -eq 1) {
                        return @() # No storage initially
                    } else {
                        # Storage exists now because of shadow copy
                        $mockStorage = [PSCustomObject]@{
                            Volume = 'Win32_Volume.DeviceID="\\\\?\\Volume{12345678-0000-0000-0000-100000000000}\\"'
                            DiffVolume = 'Win32_Volume.DeviceID="\\\\?\\Volume{12345678-0000-0000-0000-100000000000}\\"'
                            MaxSpace = 0
                        }
                        $mockStorage | Add-Member -MemberType ScriptMethod -Name "Put" -Value {
                            $script:limitApplied = $true
                            return [PSCustomObject]@{ Path = 'Win32_ShadowStorage' }
                        }
                        return @($mockStorage)
                    }
                }
                elseif ($Class -eq 'Win32_Volume') {
                    return @(
                        [PSCustomObject]@{
                            DeviceID = '\\?\Volume{12345678-0000-0000-0000-100000000000}\'
                            DriveLetter = 'C:'
                        }
                    )
                }
            }

            Mock -CommandName New-VSSShadowCopy -MockWith {
                $script:tempShadowCreated = $true
                return [PSCustomObject]@{
                    Success = $true
                    ShadowID = '{98765432-0000-0000-0000-100000000000}'
                }
            }

            Mock -CommandName Remove-VSSShadowCopy -MockWith {
                $script:tempShadowDeleted = $true
                return [PSCustomObject]@{ Success = $true }
            }

            # WMI class mock to fail direct creation
            Mock -CommandName Get-WmiObject -ParameterFilter { $List -eq $true } -MockWith {
                $mockClass = [PSCustomObject]@{}
                $mockClass | Add-Member -MemberType ScriptMethod -Name "Create" -Value {
                    param($volume, $diffVolume, $limit)
                    # direct create fails
                    return [PSCustomObject]@{ ReturnValue = 10 }
                }
                return $mockClass
            }

            $res = Set-VSSShadowStorageLimit -VolumePath 'C:' -MaxSpaceBytes 5368709120
            $res.Success | Should Be $true
            $script:tempShadowCreated | Should Be $true
            $script:limitApplied | Should Be $true
            $script:tempShadowDeleted | Should Be $true
        }
    }
}
