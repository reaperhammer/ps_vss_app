# PowerShell helpers for Volume Shadow Copy Service (VSS) operations on Windows.
# Requires Windows PowerShell 5.1 for WMI and WPF compatibility.

$DebugPreference = "SilentlyContinue"

function Test-VSSAdministrator {
    <#
    .SYNOPSIS
        Returns whether the current PowerShell process is elevated.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        Write-Verbose "Unable to determine administrator status: $($_.Exception.Message)"
        return $false
    }
}

function ConvertTo-VSSCanonicalVolumeName {
    <#
    .SYNOPSIS
        Normalizes a volume name/device ID for comparison.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [AllowNull()]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }

    return (($Path.Trim() -replace '\\+$', '').ToUpperInvariant())
}

function Get-VSSUsageColor {
    <#
    .SYNOPSIS
        Returns the UI color used for a volume usage percentage.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [double]$UsagePercent
    )

    if ($UsagePercent -lt 70) {
        return "#4CAF50"
    }
    elseif ($UsagePercent -lt 85) {
        return "#FF9800"
    }

    return "#F44336"
}

function Get-VSSShadowStateName {
    <#
    .SYNOPSIS
        Converts Win32_ShadowCopy numeric or textual state values into display text.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [AllowNull()]
        $State
    )

    if ($null -eq $State) {
        return "Unknown"
    }

    $stateMap = @{
        0  = "Unknown"
        1  = "Preparing"
        2  = "ProcessingPrepare"
        3  = "Prepared"
        4  = "ProcessingPreCommit"
        5  = "PreCommitted"
        6  = "ProcessingCommit"
        7  = "Committed"
        8  = "ProcessingPostCommit"
        9  = "ProcessingPreFinalCommit"
        10 = "PreFinalCommitted"
        11 = "ProcessingPostFinalCommit"
        12 = "Created"
        13 = "Aborted"
        14 = "Deleted"
        15 = "PostCommitted"
        16 = "Count"
    }

    $stateCode = 0
    if ([int]::TryParse([string]$State, [ref]$stateCode)) {
        if ($stateMap.ContainsKey($stateCode)) {
            return $stateMap[$stateCode]
        }

        return "Unknown ($stateCode)"
    }

    $stateText = ([string]$State).Trim()
    if ([string]::IsNullOrWhiteSpace($stateText)) {
        return "Unknown"
    }

    # Preserve recognizable WMI/provider strings while normalizing common variants.
    switch -Regex ($stateText) {
        '^Created$' { return "Created" }
        '^Active$' { return "Active" }
        '^Preparing$' { return "Preparing" }
        '^Failed$' { return "Failed" }
        '^Committed$' { return "Committed" }
        '^Deleted$' { return "Deleted" }
        '^Aborted$' { return "Aborted" }
        default { return $stateText }
    }
}

function Get-VSSShadowStateColor {
    <#
    .SYNOPSIS
        Returns the UI color used for a shadow copy state.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [AllowNull()]
        $State
    )

    $stateName = Get-VSSShadowStateName -State $State

    switch -Regex ($stateName) {
        '^(Created|Committed|PostCommitted)$' { return "#4CAF50" }
        '^(Active|Processing.*|Pre.*|Prepared)$' { return "#2196F3" }
        '^Preparing$' { return "#FF9800" }
        '^(Failed|Aborted|Deleted)$' { return "#F44336" }
        default { return "#757575" }
    }
}

function ConvertFrom-VSSWmiDateTime {
    <#
    .SYNOPSIS
        Converts a WMI datetime string into a local DateTime value.
    #>
    [CmdletBinding()]
    [OutputType([Nullable[datetime]])]
    param(
        [AllowNull()]
        [string]$WmiDateTime
    )

    if ([string]::IsNullOrWhiteSpace($WmiDateTime)) {
        return $null
    }

    try {
        return [System.Management.ManagementDateTimeConverter]::ToDateTime($WmiDateTime).ToLocalTime()
    }
    catch {
        try {
            return [datetime]$WmiDateTime
        }
        catch {
            Write-Verbose "Unable to convert WMI datetime '$WmiDateTime': $($_.Exception.Message)"
            return $null
        }
    }
}

function Get-VSSAgeText {
    <#
    .SYNOPSIS
        Returns a compact human-readable age for a DateTime.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [AllowNull()]
        $DateTime
    )

    if ($null -eq $DateTime) {
        return "Unknown"
    }

    try {
        $created = [datetime]$DateTime
        $timeSpan = (Get-Date) - $created

        if ($timeSpan.TotalDays -ge 1) {
            return "$([math]::Floor($timeSpan.TotalDays))d"
        }
        elseif ($timeSpan.TotalHours -ge 1) {
            return "$([math]::Floor($timeSpan.TotalHours))h"
        }
        elseif ($timeSpan.TotalMinutes -ge 1) {
            return "$([math]::Floor($timeSpan.TotalMinutes))m"
        }

        return "<1m"
    }
    catch {
        return "Unknown"
    }
}

function Resolve-VSSVolumePath {
    <#
    .SYNOPSIS
        Resolves a drive letter or Win32_Volume device ID into a consistent volume object.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$VolumePath
    )

    $inputPath = $VolumePath.Trim()
    $driveLetter = $null

    if ($inputPath -match '^[A-Za-z]:?\\?$') {
        $driveLetter = $inputPath.Substring(0, 1).ToUpperInvariant() + ":"
    }

    $volumes = @(Get-WmiObject -Class Win32_Volume -ErrorAction Stop)
    $volume = $null

    if ($driveLetter) {
        $volume = $volumes | Where-Object { $_.DriveLetter -eq $driveLetter } | Select-Object -First 1
    }
    else {
        $canonicalInput = ConvertTo-VSSCanonicalVolumeName -Path $inputPath
        $volume = $volumes | Where-Object {
            (ConvertTo-VSSCanonicalVolumeName -Path $_.DeviceID) -eq $canonicalInput -or
            (ConvertTo-VSSCanonicalVolumeName -Path $_.Name) -eq $canonicalInput
        } | Select-Object -First 1
    }

    if (-not $volume) {
        throw "No VSS-capable volume found for '$VolumePath'. Use a drive letter such as C: or a Win32_Volume DeviceID."
    }

    [PSCustomObject]@{
        PSTypeName      = "VSS.ResolvedVolume"
        InputPath       = $VolumePath
        DriveLetter     = $volume.DriveLetter
        DeviceID        = $volume.DeviceID
        CanonicalDevice = ConvertTo-VSSCanonicalVolumeName -Path $volume.DeviceID
        Name            = $volume.Name
        Label           = $volume.Label
        FileSystem      = $volume.FileSystem
        Volume          = $volume
    }
}

function Get-VSSSupportedVolumes {
    <#
    .SYNOPSIS
        Lists fixed volumes that support Volume Shadow Copy Service operations.
    .DESCRIPTION
        Returns structured objects instead of writing formatted text, so the function can be used by scripts and the GUI.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    try {
        $volumes = Get-WmiObject -Class Win32_Volume -ErrorAction Stop |
            Where-Object { $_.DriveType -eq 3 -and $_.DriveLetter }

        foreach ($volume in $volumes) {
            $capacityBytes = if ($null -ne $volume.Capacity) { [int64]$volume.Capacity } else { 0 }
            $freeSpaceBytes = if ($null -ne $volume.FreeSpace) { [int64]$volume.FreeSpace } else { 0 }
            $usedSpaceBytes = [math]::Max([int64]0, $capacityBytes - $freeSpaceBytes)
            $usagePercent = if ($capacityBytes -gt 0) { [math]::Round(($usedSpaceBytes / $capacityBytes) * 100, 1) } else { 0 }

            [PSCustomObject]@{
                PSTypeName       = "VSS.Volume"
                DriveLetter      = $volume.DriveLetter
                DeviceID         = $volume.DeviceID
                Name             = $volume.Name
                VolumeName       = if ($volume.Label) { $volume.Label } else { "Local Disk" }
                Label            = $volume.Label
                FileSystem       = $volume.FileSystem
                DriveType        = $volume.DriveType
                CapacityBytes    = $capacityBytes
                FreeSpaceBytes   = $freeSpaceBytes
                UsedSpaceBytes   = $usedSpaceBytes
                CapacityGB       = [math]::Round($capacityBytes / 1GB, 2)
                FreeSpaceGB      = [math]::Round($freeSpaceBytes / 1GB, 2)
                UsedSpaceGB      = [math]::Round($usedSpaceBytes / 1GB, 2)
                UsagePercent     = $usagePercent
                UsagePercentText = "$usagePercent%"
                UsageColor       = Get-VSSUsageColor -UsagePercent $usagePercent
            }
        }
    }
    catch {
        Write-Error "Failed to retrieve VSS-supported volumes: $($_.Exception.Message)"
    }
}

function Get-VSSShadowCopies {
    <#
    .SYNOPSIS
        Lists shadow copies for a specific volume.
    .DESCRIPTION
        Accepts a drive letter or Win32_Volume DeviceID and returns structured shadow copy objects.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$VolumePath
    )

    try {
        $resolvedVolume = Resolve-VSSVolumePath -VolumePath $VolumePath -ErrorAction Stop
        $canonicalDeviceId = $resolvedVolume.CanonicalDevice

        $shadowCopies = Get-WmiObject -Class Win32_ShadowCopy -ErrorAction Stop |
            Where-Object { (ConvertTo-VSSCanonicalVolumeName -Path $_.VolumeName) -eq $canonicalDeviceId }

        foreach ($shadow in $shadowCopies) {
            $created = ConvertFrom-VSSWmiDateTime -WmiDateTime $shadow.InstallDate
            $stateName = Get-VSSShadowStateName -State $shadow.State
            $sizeBytes = if ($shadow.PSObject.Properties["Size"] -and $null -ne $shadow.Size) { [int64]$shadow.Size } else { $null }

            [PSCustomObject]@{
                PSTypeName       = "VSS.ShadowCopy"
                ShadowID         = $shadow.ID
                ID               = $shadow.ID
                DeviceObject     = $shadow.DeviceObject
                VolumeName       = $shadow.VolumeName
                VolumePath       = $shadow.VolumeName
                RequestedVolume  = $VolumePath
                CreationTime     = $created
                CreationTimeText = if ($created) { $created.ToString("yyyy-MM-dd HH:mm:ss") } else { "Unknown" }
                InstallDate      = $shadow.InstallDate
                State            = $stateName
                StateCode        = $shadow.State
                StateColor       = Get-VSSShadowStateColor -State $shadow.State
                Persistent       = $shadow.Persistent
                ClientAccessible = $shadow.ClientAccessible
                NoWriters        = $shadow.NoWriters
                SizeBytes        = $sizeBytes
                SizeMB           = if ($null -ne $sizeBytes) { [math]::Round($sizeBytes / 1MB, 1) } else { "N/A" }
                Age              = Get-VSSAgeText -DateTime $created
                RawObject        = $shadow
            }
        }
    }
    catch {
        Write-Error "Failed to list shadow copies for volume '$VolumePath': $($_.Exception.Message)"
    }
}

function New-VSSShadowCopy {
    <#
    .SYNOPSIS
        Creates a new shadow copy for a specified volume.
    .DESCRIPTION
        Creates a new Volume Shadow Copy and returns a structured result object. Requires administrator privileges.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$VolumePath,

        [string]$Description = "PowerShell Created Shadow Copy",

        [string]$Context = "ClientAccessible"
    )

    try {
        if (-not (Test-VSSAdministrator)) {
            throw "Administrator privileges are required to create VSS shadow copies."
        }

        $resolvedVolume = Resolve-VSSVolumePath -VolumePath $VolumePath -ErrorAction Stop

        if ($PSCmdlet.ShouldProcess($resolvedVolume.DeviceID, "Create VSS shadow copy")) {
            Write-Verbose "Creating shadow copy for $($resolvedVolume.DeviceID). Description: $Description"

            $shadowClass = Get-WmiObject -List Win32_ShadowCopy -ErrorAction Stop
            $result = $shadowClass.Create($resolvedVolume.DeviceID, $Context)
            $returnValue = if ($result.PSObject.Properties["ReturnValue"]) { [int]$result.ReturnValue } else { $null }
            $shadowId = if ($result.PSObject.Properties["ShadowID"]) { $result.ShadowID } else { $null }
            $success = ($returnValue -eq 0 -and -not [string]::IsNullOrWhiteSpace($shadowId))

            $response = [PSCustomObject]@{
                PSTypeName  = "VSS.CreateResult"
                Success     = $success
                ReturnValue = $returnValue
                ShadowID    = $shadowId
                VolumePath  = $VolumePath
                DeviceID    = $resolvedVolume.DeviceID
                Context     = $Context
                Description = $Description
            }

            if (-not $success) {
                Write-Error "Failed to create shadow copy for '$VolumePath'. WMI return code: $returnValue"
            }

            return $response
        }
    }
    catch {
        Write-Error "Failed to create shadow copy for volume '$VolumePath': $($_.Exception.Message)"
    }
}

function Remove-VSSShadowCopy {
    <#
    .SYNOPSIS
        Removes shadow copies for a specific volume.
    .DESCRIPTION
        Deletes either a specific shadow copy or all shadow copies for a specified volume. Uses PowerShell -Confirm/-WhatIf safety semantics.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$VolumePath,

        [string]$ShadowCopyID
    )

    try {
        if (-not (Test-VSSAdministrator)) {
            throw "Administrator privileges are required to delete VSS shadow copies."
        }

        $resolvedVolume = Resolve-VSSVolumePath -VolumePath $VolumePath -ErrorAction Stop
        $canonicalDeviceId = $resolvedVolume.CanonicalDevice

        $shadowCopies = @(Get-WmiObject -Class Win32_ShadowCopy -ErrorAction Stop |
            Where-Object { (ConvertTo-VSSCanonicalVolumeName -Path $_.VolumeName) -eq $canonicalDeviceId })

        if ($ShadowCopyID) {
            $shadowCopies = @($shadowCopies | Where-Object { $_.ID -eq $ShadowCopyID })
        }

        if (-not $shadowCopies -or $shadowCopies.Count -eq 0) {
            Write-Verbose "No matching shadow copies found for volume '$VolumePath'."
            return
        }

        foreach ($shadow in $shadowCopies) {
            $target = "$($shadow.ID) on $($resolvedVolume.DeviceID)"

            if ($PSCmdlet.ShouldProcess($target, "Delete VSS shadow copy")) {
                $result = $shadow.Delete()
                $returnValue = if ($result -and $result.PSObject.Properties["ReturnValue"]) { [int]$result.ReturnValue } elseif ($null -eq $result) { 0 } else { [int]$result }
                $success = ($returnValue -eq 0)

                $response = [PSCustomObject]@{
                    PSTypeName  = "VSS.DeleteResult"
                    Success     = $success
                    ReturnValue = $returnValue
                    ShadowID    = $shadow.ID
                    VolumePath  = $VolumePath
                    DeviceID    = $resolvedVolume.DeviceID
                }

                if (-not $success) {
                    Write-Error "Failed to delete shadow copy '$($shadow.ID)'. WMI return code: $returnValue"
                }

                $response
            }
        }
    }
    catch {
        Write-Error "Failed to delete shadow copies for volume '$VolumePath': $($_.Exception.Message)"
    }
}

function Get-VSSShadowStorage {
    <#
    .SYNOPSIS
        Queries shadow storage associations (diff area limits) for volumes.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$VolumePath
    )

    try {
        if (-not (Test-VSSAdministrator)) {
            throw "Administrator privileges are required to query VSS shadow storage."
        }

        # Get list of all supported volumes for mapping reference ID -> Drive Letter
        $allvols = @(Get-VSSSupportedVolumes)
        
        $canonicalDeviceId = $null
        if ($VolumePath) {
            $resolvedVolume = Resolve-VSSVolumePath -VolumePath $VolumePath -ErrorAction Stop
            $canonicalDeviceId = $resolvedVolume.CanonicalDevice
        }

        $storages = @(Get-WmiObject -Class Win32_ShadowStorage -ErrorAction Stop)

        foreach ($storage in $storages) {
            # Extract DeviceID from Volume and DiffVolume reference paths
            $volId = $null
            $diffVolId = $null
            if ($storage.Volume -match 'DeviceID="(?<id>[^"]+)"') {
                $volId = $Matches.id -replace '\\+$', ''
            }
            if ($storage.DiffVolume -match 'DeviceID="(?<id>[^"]+)"') {
                $diffVolId = $Matches.id -replace '\\+$', ''
            }

            $canonicalVolId = if ($volId) { $volId.ToUpperInvariant() } else { $null }
            
            # Filter if a specific VolumePath was requested
            if ($canonicalDeviceId -and $canonicalVolId -ne $canonicalDeviceId) {
                continue
            }

            # Map to drive letters
            $volInfo = $allvols | Where-Object { (ConvertTo-VSSCanonicalVolumeName -Path $_.DeviceID) -eq $canonicalVolId }
            $diffVolInfo = $allvols | Where-Object { (ConvertTo-VSSCanonicalVolumeName -Path $_.DeviceID) -eq ($diffVolId.ToUpperInvariant()) }

            $volLetter = if ($volInfo) { $volInfo.DriveLetter } else { $volId }
            $diffVolLetter = if ($diffVolInfo) { $diffVolInfo.DriveLetter } else { $diffVolId }

            [PSCustomObject]@{
                PSTypeName       = "VSS.ShadowStorage"
                Volume           = $volId
                DiffVolume       = $diffVolId
                VolumeLetter     = $volLetter
                DiffVolumeLetter = $diffVolLetter
                MaxSpace         = $storage.MaxSpace
                AllocatedSpace   = $storage.AllocatedSpace
                UsedSpace        = $storage.UsedSpace
                MaxSpaceGB       = if ($storage.MaxSpace -eq [UInt64]::MaxValue -or $storage.MaxSpace -eq 0 -or $storage.MaxSpace -eq [Int64]::MaxValue -or $storage.MaxSpace -eq 9223372036854775807) { "UNLIMITED" } else { [math]::Round($storage.MaxSpace / 1GB, 2) }
                AllocatedSpaceGB = [math]::Round($storage.AllocatedSpace / 1GB, 2)
                UsedSpaceGB      = [math]::Round($storage.UsedSpace / 1GB, 2)
                StorageObject    = $storage
            }
        }
    }
    catch {
        Write-Error "Failed to retrieve shadow storage: $($_.Exception.Message)"
    }
}

function Set-VSSShadowStorageLimit {
    <#
    .SYNOPSIS
        Sets the shadow storage limit (MaxSpace) for a volume. Creates association if none exists.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$VolumePath,

        [string]$DiffVolumePath,

        [Parameter(Mandatory = $true)]
        [int64]$MaxSpaceBytes
    )

    try {
        if (-not (Test-VSSAdministrator)) {
            throw "Administrator privileges are required to configure VSS shadow storage."
        }

        $resolvedVol = Resolve-VSSVolumePath -VolumePath $VolumePath -ErrorAction Stop
        $resolvedDiff = if ($DiffVolumePath) { Resolve-VSSVolumePath -VolumePath $DiffVolumePath -ErrorAction Stop } else { $resolvedVol }
        
        $canonicalVol = $resolvedVol.CanonicalDevice
        $limit = if ($MaxSpaceBytes -lt 0) { [UInt64]::MaxValue } else { [UInt64]$MaxSpaceBytes }

        $storage = Get-WmiObject -Class Win32_ShadowStorage | Where-Object {
            $volId = $null
            if ($_.Volume -match 'DeviceID="(?<id>[^"]+)"') { $volId = $Matches.id.ToUpperInvariant() }
            $volId -eq $canonicalVol
        }

        $success = $false
        if ($storage) {
            # Modify existing MaxSpace
            $storage.MaxSpace = $limit
            $result = $storage.Put()
            $success = ($null -ne $result)
        } else {
            # Create new association
            $volRef = "Win32_Volume.DeviceID=""$($resolvedVol.DeviceID)"""
            $diffRef = "Win32_Volume.DeviceID=""$($resolvedDiff.DeviceID)"""
            
            try {
                $class = [wmiclass]"root\cimv2:Win32_ShadowStorage"
                $result = $class.Create($volRef, $diffRef, $limit)
                $returnValue = if ($result -and $result.PSObject.Properties["ReturnValue"]) { [int]$result.ReturnValue } else { [int]$result }
                $success = ($returnValue -eq 0)
            } catch {
                # Fallback to vssadmin if WMI fails or complains about method signatures
                $volLetter = $resolvedVol.DriveLetter
                $diffVolLetter = $resolvedDiff.DriveLetter
                $limitStr = if ($MaxSpaceBytes -lt 0) { "UNLIMITED" } else { "$($MaxSpaceBytes)" }
                $cmdResult = cmd.exe /c vssadmin add shadowstorage /For=$volLetter /On=$diffVolLetter /Max=$limitStr 2>&1
                $success = ($LASTEXITCODE -eq 0)
                if (-not $success) {
                     throw "Failed to create shadow storage via WMI and vssadmin add: $cmdResult"
                }
            }
        }

        [PSCustomObject]@{
            Success    = $success
            VolumePath = $VolumePath
            MaxSpace   = $limit
        }
    }
    catch {
        Write-Error "Failed to configure shadow storage for volume '$VolumePath': $($_.Exception.Message)"
    }
}

function Get-VSSWriters {
    <#
    .SYNOPSIS
        Queries the current list of VSS Writers and their states.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    try {
        $vssadminPath = Join-Path $env:SystemRoot "System32\vssadmin.exe"
        if (-not (Test-Path $vssadminPath)) {
            throw "vssadmin.exe not found on this system. Unable to query VSS Writers."
        }

        # Run vssadmin list writers and grab output
        $output = & vssadmin list writers 2>&1
        $list = @()
        $current = $null

        foreach ($line in $output) {
            $trimmed = $line.Trim()
            if ($trimmed -match "Writer name:\s*'(?<Name>[^']+)'") {
                if ($null -ne $current) {
                    $list += [PSCustomObject]$current
                }
                $current = @{
                    PSTypeName = "VSS.Writer"
                    WriterName = $Matches.Name
                    WriterId   = "Unknown"
                    State      = "Unknown"
                    StateCode  = 0
                    LastError  = "Unknown"
                    StatusColor = "#F44336" # Default red
                }
            }
            elseif ($trimmed -match "Writer Id:\s*{(?<Id>[^}]+)}") {
                $current.WriterId = $Matches.Id
            }
            elseif ($trimmed -match "State:\s*\[(?<Code>\d+)\]\s*(?<StateText>.*)") {
                $current.StateCode = [int]$Matches.Code
                $current.State = $Matches.StateText.Trim()
                # Status coloring
                if ($current.StateCode -eq 1) {
                    $current.StatusColor = "#4CAF50" # Stable: green
                } else {
                    $current.StatusColor = "#FF9800" # Changing/Freeze: orange
                }
            }
            elseif ($trimmed -match "Last error:\s*(?<Error>.*)") {
                $current.LastError = $Matches.Error.Trim()
                if ($current.LastError -ne "No error") {
                    $current.StatusColor = "#F44336" # Red if there's a last error
                }
            }
        }

        if ($null -ne $current) {
            $list += [PSCustomObject]$current
        }

        $list
    }
    catch {
        Write-Error "Failed to retrieve VSS Writers: $($_.Exception.Message)"
    }
}

function Mount-VSSShadowCopy {
    <#
    .SYNOPSIS
        Mounts a VSS shadow copy as a folder symlink.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ShadowCopyID,

        [Parameter(Mandatory = $true)]
        [string]$MountPath
    )

    try {
        if (-not (Test-VSSAdministrator)) {
            throw "Administrator privileges are required to mount VSS shadow copies."
        }

        $shadow = Get-WmiObject -Class Win32_ShadowCopy | Where-Object { $_.ID -eq $ShadowCopyID }
        if (-not $shadow) {
            throw "Shadow copy with ID '$ShadowCopyID' not found."
        }

        $deviceObject = $shadow.DeviceObject
        if (-not $deviceObject) {
            throw "Shadow copy does not expose a valid DeviceObject path."
        }

        $target = $deviceObject + "\"

        # Create parent directory if needed
        $parent = Split-Path -Path $MountPath -Parent
        if ($parent -and -not (Test-Path -Path $parent)) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }

        # Clear existing link or folder
        if (Test-Path -Path $MountPath) {
            cmd.exe /c rmdir "$MountPath" 2>&1 | Out-Null
            Remove-Item -Path $MountPath -Force -Recurse -ErrorAction SilentlyContinue | Out-Null
        }

        # Use cmd mklink /d since PowerShell's New-Item validates target paths, which fails for raw VSS devices
        $cmdResult = cmd.exe /c mklink /d "$MountPath" "$target" 2>&1
        $success = ($LASTEXITCODE -eq 0)

        if (-not $success) {
            throw "mklink failed: $cmdResult"
        }

        [PSCustomObject]@{
            Success   = $true
            MountPath = $MountPath
            Target    = $target
            ShadowID  = $ShadowCopyID
        }
    }
    catch {
        Write-Error "Failed to mount shadow copy: $($_.Exception.Message)"
    }
}

function Dismount-VSSShadowCopy {
    <#
    .SYNOPSIS
        Safely dismounts a VSS shadow copy folder symlink.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MountPath
    )

    try {
        $success = $false
        if (Test-Path -Path $MountPath) {
            # rmdir is safe for junctions/symlinks and doesn't delete target files
            $result = cmd.exe /c rmdir "$MountPath" 2>&1
            $success = ($LASTEXITCODE -eq 0)
            if (-not $success) {
                throw "rmdir failed: $result"
            }
        } else {
            $success = $true
        }

        [PSCustomObject]@{
            Success   = $success
            MountPath = $MountPath
        }
    }
    catch {
        Write-Error "Failed to dismount shadow copy at '$MountPath': $($_.Exception.Message)"
    }
}


# Example usage:
# Get-VSSSupportedVolumes
# Get-VSSShadowCopies -VolumePath "E:\"
# New-VSSShadowCopy -VolumePath "C:\" -Description "Test Shadow Copy"
# Remove-VSSShadowCopy -VolumePath "E:\" -Confirm:$false
# Remove-VSSShadowCopy -VolumePath "E:\" -WhatIf
