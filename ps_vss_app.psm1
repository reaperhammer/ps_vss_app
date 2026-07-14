# ps_vss_app - PowerShell module for Volume Shadow Copy Service (VSS) management on Windows.
# Requires Windows PowerShell 5.1 (Desktop edition) for WMI, WPF, and dynamic C# compilation.
# This module exposes VSS helper functions used by both interactive commands and the VSSManager.ps1 GUI.
#
# Usage:
#   Import-Module ./ps_vss_app.psd1
#   Get-VSSSupportedVolumes | Format-Table
#   New-VSSShadowCopy -VolumePath 'C:' -Context Backup
#
# The module is the canonical home of these functions. The legacy main.ps1 file is now a
# backward-compatibility shim that simply imports this module.

#Requires -Version 5.1

Set-StrictMode -Version 2.0

# Make the module's own directory discoverable to internal helpers that need to locate
# VssBackupHelper.cs and bin\VssBackupHelper.exe. In a module, $PSScriptRoot inside a
# function body resolves to the module folder, but caching it here also makes it
# accessible from the module scope itself.
$script:ModuleRoot = $PSScriptRoot

function Test-VSSAdministrator {
    <#
    .SYNOPSIS
        Returns whether the current PowerShell process is elevated.
    .DESCRIPTION
        Checks the current process token against the local Administrators group via
        the WindowsPrincipal role API. VSS operations (creating or deleting shadow
        copies, configuring shadow storage, querying writers) all require elevation,
        and most of this module's public functions call this helper at the top to
        fail fast with a clear error rather than letting WMI / VSS surface vague
        access-denied messages later.
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
    .DESCRIPTION
        Returns a canonical form of a volume device path by collapsing repeated
        backslashes, uppercasing the result, and trimming trailing separators.
        Used internally to compare WMI-reported device IDs (which escape backslashes
        in their string representation) against paths supplied by callers. Returns
        $null for blank or null input.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [AllowNull()]
        [string]$Path
    )

    process {
        if ([string]::IsNullOrWhiteSpace($Path)) {
            return $null
        }

        $normalized = $Path.Trim() -replace '\\+', '\'
        if ($normalized.StartsWith('\')) {
            $normalized = '\' + $normalized
        }
        return ($normalized -replace '\\+$', '').ToUpperInvariant()
    }
}

function Get-VSSUsageColor {
    <#
    .SYNOPSIS
        Returns the UI color used for a volume usage percentage.
    .DESCRIPTION
        Maps a 0-100 usage percentage to a hex color string suitable for WPF brushes.
        Green below 70%, orange 70-85%, red 85% and above. The thresholds are aligned
        with the WPF GUI's volume-management tab so script consumers and the GUI
        stay in sync.
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

function Invoke-NativeCommand {
    <#
    .SYNOPSIS
        Private helper that invokes an external command and captures its output.
    .DESCRIPTION
        Wraps `&` so tests can intercept native command invocations via Pester's
        Mock -CommandName. PowerShell's built-in `&` operator isn't mockable, but a
        function is.
    #>
    [CmdletBinding()]
    [OutputType([object[]])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CommandPath,

        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    & $CommandPath @Arguments 2>&1
}

function Get-VSSShadowStateName {
    <#
    .SYNOPSIS
        Converts Win32_ShadowCopy numeric or textual state values into display text.
    .DESCRIPTION
        WMI reports shadow copy state as an integer code in some locales and as a
        string like "Created" or "Committed" in others. This function accepts either
        representation and returns a human-readable English label. Unknown values
        pass through as "Unknown (<code>)" or the original text so callers can
        still display something.
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

function Invoke-NativeCommand {
    <#
    .SYNOPSIS
        Private helper that invokes an external command and captures its output.
    .DESCRIPTION
        Wraps `&` so tests can intercept native command invocations via Pester's
        Mock -CommandName. PowerShell's built-in `&` operator isn't mockable, but a
        function is.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CommandPath,

        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    & $CommandPath @Arguments 2>&1
}

function Get-VSSShadowStateColor {
    <#
    .SYNOPSIS
        Returns the UI color used for a shadow copy state.
    .DESCRIPTION
        Maps a shadow copy's state (as accepted by Get-VSSShadowStateName) to a
        hex color string for the WPF GUI. Green for terminal success states
        (Created, Committed), blue for in-progress states, orange for transient
        states, red for failed / aborted / deleted, gray for unknown.
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
    .DESCRIPTION
        WMI's ManagementDateTimeConverter.ToDateTime returns UTC. This function
        calls that converter and converts to local time, with a fallback to
        the built-in [datetime] cast for non-WMI inputs. Returns $null for
        blank or unparseable input.
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
    .DESCRIPTION
        Formats the time elapsed since -DateTime as "<n>d", "<n>h", "<n>m" or
        "<1m" for sub-minute intervals. Used by the WPF GUI's Shadow Copies tab
        so users can see at a glance how old each snapshot is. Returns "Unknown"
        for null or unparseable input.
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
    .DESCRIPTION
        Returns a PSCustomObject (VSS.ResolvedVolume) describing the volume. The raw
        Win32_Volume WMI object is included as the OriginalWmiObject property for
        callers that need additional WMI properties not surfaced on the canonical shape.
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
        OriginalWmiObject = $volume
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

function Get-VSSBackupHelperPath {
    <#
    .SYNOPSIS
        Returns the path to VssBackupHelper.exe, compiling from VssBackupHelper.cs on the fly if needed.
    .DESCRIPTION
        Resolution order:
          1. Pre-built bin\VssBackupHelper.exe in the module directory.
          2. Compile VssBackupHelper.cs with the in-box .NET Framework 4.x csc.exe (64-bit, then 32-bit).
          3. Compile VssBackupHelper.cs with the Roslyn-based fallback (System.CodeDom.Compiler via Add-Type).
        Returns $null if no path could be produced.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    $moduleRoot = $script:ModuleRoot
    if (-not $moduleRoot) {
        $moduleRoot = $PSScriptRoot
    }

    $helperPath = Join-Path $moduleRoot "bin\VssBackupHelper.exe"
    if (Test-Path $helperPath) {
        return $helperPath
    }

    $sourcePath = Join-Path $moduleRoot "VssBackupHelper.cs"
    if (-not (Test-Path $sourcePath)) {
        return $null
    }

    $binDir = Join-Path $moduleRoot "bin"
    if (-not (Test-Path $binDir)) {
        New-Item -ItemType Directory -Path $binDir -Force | Out-Null
    }

    # Strategy 1: csc.exe from the .NET Framework SDK
    $cscCandidates = @(
        (Join-Path $env:SystemRoot "Microsoft.NET\Framework64\v4.0.30319\csc.exe")
        (Join-Path $env:SystemRoot "Microsoft.NET\Framework\v4.0.30319\csc.exe")
    )
    foreach ($cscPath in $cscCandidates) {
        if (Test-Path $cscPath) {
            Write-Verbose "Compiling VssBackupHelper.cs on the fly via $cscPath..."
            try {
                & $cscPath /target:exe /out:$helperPath $sourcePath | Out-Null
                if (Test-Path $helperPath) {
                    return $helperPath
                }
            }
            catch {
                Write-Verbose "csc.exe compilation failed at ${cscPath}: $($_.Exception.Message)"
            }
        }
    }

    # Strategy 2: Roslyn fallback via System.CodeDom.Compiler (works on any system with .NET 4.x
    # runtime; the GAC'd Microsoft.CSharp and System.CodeDom assemblies are present in .NET 4.x).
    Write-Verbose "csc.exe unavailable; attempting Roslyn fallback via System.CodeDom.Compiler..."
    try {
        Add-Type -AssemblyName System.CodeDom
        Add-Type -AssemblyName Microsoft.CSharp
        $compilerParameters = New-Object System.CodeDom.Compiler.CompilerParameters
        $compilerParameters.OutputAssembly = $helperPath
        $compilerParameters.GenerateExecutable = $true
        $compilerParameters.GenerateInMemory = $false
        $compilerParameters.TreatWarningsAsErrors = $false
        $compilerParameters.IncludeDebugInformation = $false
        $compilerParameters.CompilerOptions = "/nologo /optimize"

        $codeDomProvider = [System.CodeDom.Compiler.CodeDomProvider]::CreateProvider("CSharp", [System.CodeDom.Compiler.CompilerParameters]::new().ReferencedAssemblies)
        # CreateProvider above may not preserve our parameters; use the overload that accepts parameters.
        $codeDomProvider = New-Object Microsoft.CSharp.CSharpCodeProvider
        $sourceFiles = New-Object System.Collections.Generic.List[string]
        $sourceFiles.Add($sourcePath)
        $compilerResults = $codeDomProvider.CompileAssemblyFromFile($compilerParameters, $sourceFiles)

        if ($compilerResults.Errors.Count -gt 0) {
            $messages = $compilerResults.Errors | ForEach-Object { "$($_.ErrorNumber): $($_.ErrorText)" }
            Write-Verbose "Roslyn compilation failed: $($messages -join '; ')"
        }
        elseif (Test-Path $helperPath) {
            return $helperPath
        }
    }
    catch {
        Write-Verbose "Roslyn fallback failed: $($_.Exception.Message)"
    }

    return $null
}

function New-VSSShadowCopy {
    <#
    .SYNOPSIS
        Creates a new shadow copy for a specified volume.
    .DESCRIPTION
        Creates a new Volume Shadow Copy and returns a structured result object. Requires administrator privileges.
        Selects the most appropriate creation strategy (VssBackupHelper, diskshadow, or WMI) based on
        the requested Context and what the host system supports.
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

            $strategy = Get-VSSCreationStrategy -Context $Context
            switch ($strategy) {
                'Helper' {
                    $result = New-VSSShadowCopyViaHelper -ResolvedVolume $resolvedVolume -Context $Context
                }
                'DiskShadow' {
                    $result = New-VSSShadowCopyViaDiskShadow -ResolvedVolume $resolvedVolume -Context $Context
                }
                default {
                    $result = New-VSSShadowCopyViaWmi -ResolvedVolume $resolvedVolume -Context $Context
                }
            }

            if (-not $result.Success) {
                Write-Error "Failed to create shadow copy for '$VolumePath'. $($result.ErrorDescription)"
            }

            return $result
        }
    }
    catch {
        Write-Error "Failed to create shadow copy for volume '$VolumePath': $($_.Exception.Message)"
    }
}

function Get-VSSCreationStrategy {
    <#
    .SYNOPSIS
        Internal helper that returns a [string] identifying which strategy to use.
    .DESCRIPTION
        Returns one of 'Helper', 'DiskShadow', or 'Wmi'. Public callers should always use
        New-VSSShadowCopy; this function exists to allow each strategy to be unit-tested
        in isolation.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    if ($Context -ne "ClientAccessible") {
        $helperPath = Get-VSSBackupHelperPath
        if ($helperPath) {
            return 'Helper'
        }
    }

    if ($Context -eq "Backup") {
        $diskshadowPath = Join-Path $env:SystemRoot "System32\diskshadow.exe"
        if (Test-Path $diskshadowPath) {
            return 'DiskShadow'
        }
    }

    return 'Wmi'
}

function New-VSSShadowCopyViaHelper {
    <#
    .SYNOPSIS
        Creates a shadow copy using the VssBackupHelper C# COM Interop requester.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$ResolvedVolume,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    $volLetter = $ResolvedVolume.DriveLetter
    if (-not $volLetter) {
        $volLetter = $ResolvedVolume.Name
    }

    $HelperPath = Get-VSSBackupHelperPath
    if (-not $HelperPath) {
        return [PSCustomObject]@{
            PSTypeName       = "VSS.CreateResult"
            Success          = $false
            ReturnValue      = -3
            ShadowID         = $null
            VolumePath       = $ResolvedVolume.InputPath
            DeviceID         = $ResolvedVolume.DeviceID
            Context          = $Context
            Description      = ""
            ErrorDescription = "VssBackupHelper was selected as the strategy but the helper executable is not available."
        }
    }

    try {
        $cmdResult = Invoke-NativeCommand -CommandPath $HelperPath -Arguments @($volLetter, $Context)
        $exitCode = $LASTEXITCODE

        $isSuccess = $false
        $shadowId = $null
        foreach ($line in $cmdResult) {
            # Progress markers emitted by the C# helper: "PROGRESS:<percent>:<message>".
            # Emit these to the verbose stream so the GUI's status bar can pick them up
            # via a dispatcher hook.
            if ($line -match '^PROGRESS:(?<pct>\d+):(?<msg>.+)$') {
                Write-Verbose ("VSS progress: {0}% - {1}" -f $Matches.pct, $Matches.msg)
                continue
            }
            if ($line -eq "SUCCESS") {
                $isSuccess = $true
            }
            if ($line -match '(?i)SnapshotID:\s*(?<id>{[a-f0-9\-]+})') {
                $shadowId = $Matches.id
            }
        }

        if ($exitCode -eq 0 -and $isSuccess -and $shadowId) {
            return [PSCustomObject]@{
                PSTypeName       = "VSS.CreateResult"
                Success          = $true
                ReturnValue      = 0
                ShadowID         = $shadowId
                VolumePath       = $ResolvedVolume.InputPath
                DeviceID         = $ResolvedVolume.DeviceID
                Context          = $Context
                Description      = ""
                ErrorDescription = $null
            }
        }

        $returnValue = if ($exitCode -ne 0) { $exitCode } else { -1 }
        return [PSCustomObject]@{
            PSTypeName       = "VSS.CreateResult"
            Success          = $false
            ReturnValue      = $returnValue
            ShadowID         = $null
            VolumePath       = $ResolvedVolume.InputPath
            DeviceID         = $ResolvedVolume.DeviceID
            Context          = $Context
            Description      = ""
            ErrorDescription = "VssBackupHelper execution failed. Exit code: $exitCode. Output: " + ($cmdResult -join "`n")
        }
    }
    catch {
        return [PSCustomObject]@{
            PSTypeName       = "VSS.CreateResult"
            Success          = $false
            ReturnValue      = -2
            ShadowID         = $null
            VolumePath       = $ResolvedVolume.InputPath
            DeviceID         = $ResolvedVolume.DeviceID
            Context          = $Context
            Description      = ""
            ErrorDescription = "Failed to run VssBackupHelper: $($_.Exception.Message)"
        }
    }
}

function New-VSSShadowCopyViaDiskShadow {
    <#
    .SYNOPSIS
        Creates a shadow copy by piggybacking on diskshadow.exe.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$ResolvedVolume,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    $volLetter = $ResolvedVolume.DriveLetter
    if (-not $volLetter) {
        $volLetter = $ResolvedVolume.Name
    }
    if ($volLetter -notmatch '\\$') {
        $volLetter = $volLetter + "\"
    }

    $DiskShadowPath = Join-Path $env:SystemRoot "System32\diskshadow.exe"
    if (-not (Test-Path $DiskShadowPath)) {
        return [PSCustomObject]@{
            PSTypeName       = "VSS.CreateResult"
            Success          = $false
            ReturnValue      = -4
            ShadowID         = $null
            VolumePath       = $ResolvedVolume.InputPath
            DeviceID         = $ResolvedVolume.DeviceID
            Context          = $Context
            Description      = ""
            ErrorDescription = "diskshadow.exe was not found at $DiskShadowPath."
        }
    }

    $tempScript = [System.IO.Path]::GetTempFileName()
    $scriptContent = @(
        "set verbose on",
        "set context persistent",
        "add volume $volLetter",
        "create"
    )
    $scriptContent | Out-File -FilePath $tempScript -Encoding ASCII -Force

    try {
        $diskshadowExe = "diskshadow"
        if (-not (Get-Command $diskshadowExe -ErrorAction SilentlyContinue)) {
            $diskshadowExe = $DiskShadowPath
        }

        $cmdResult = Invoke-NativeCommand -CommandPath $diskshadowExe -Arguments @('/s', $tempScript)
        $exitCode = $LASTEXITCODE

        $shadowId = $null
        foreach ($line in $cmdResult) {
            if ($line -match '(?i)Created shadow copy\s*(?<id>{[a-f0-9\-]+})') {
                $shadowId = $Matches.id
                break
            }
        }

        if ($exitCode -eq 0 -and $shadowId) {
            return [PSCustomObject]@{
                PSTypeName       = "VSS.CreateResult"
                Success          = $true
                ReturnValue      = 0
                ShadowID         = $shadowId
                VolumePath       = $ResolvedVolume.InputPath
                DeviceID         = $ResolvedVolume.DeviceID
                Context          = $Context
                Description      = ""
                ErrorDescription = $null
            }
        }

        $returnValue = if ($exitCode -ne 0) { $exitCode } else { -1 }
        return [PSCustomObject]@{
            PSTypeName       = "VSS.CreateResult"
            Success          = $false
            ReturnValue      = $returnValue
            ShadowID         = $null
            VolumePath       = $ResolvedVolume.InputPath
            DeviceID         = $ResolvedVolume.DeviceID
            Context          = $Context
            Description      = ""
            ErrorDescription = "diskshadow execution failed. Output: " + ($cmdResult -join "`n")
        }
    }
    catch {
        return [PSCustomObject]@{
            PSTypeName       = "VSS.CreateResult"
            Success          = $false
            ReturnValue      = -2
            ShadowID         = $null
            VolumePath       = $ResolvedVolume.InputPath
            DeviceID         = $ResolvedVolume.DeviceID
            Context          = $Context
            Description      = ""
            ErrorDescription = "Failed to run diskshadow: $($_.Exception.Message)"
        }
    }
    finally {
        if (Test-Path $tempScript) {
            Remove-Item -Path $tempScript -Force -ErrorAction SilentlyContinue
        }
    }
}

function New-VSSShadowCopyViaWmi {
    <#
    .SYNOPSIS
        Creates a shadow copy via the WMI Win32_ShadowCopy.Create method (ClientAccessible fallback).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$ResolvedVolume,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    $shadowClass = Get-WmiObject -List Win32_ShadowCopy -ErrorAction Stop
    $result = $shadowClass.Create($ResolvedVolume.DeviceID, $Context)
    $returnValue = if ($result.PSObject.Properties["ReturnValue"]) { [int]$result.ReturnValue } else { $null }
    $shadowId = if ($result.PSObject.Properties["ShadowID"]) { $result.ShadowID } else { $null }
    $success = ($returnValue -eq 0 -and -not [string]::IsNullOrWhiteSpace($shadowId))

    $errorDescription = $null
    if (-not $success) {
        if ($returnValue -eq 5) {
            $errorDescription = "Unsupported shadow copy context (WMI return code 5). Note that WMI's Win32_ShadowCopy.Create method does not support the 'Backup' context on Windows Client editions (e.g., Windows 10/11), which only support the 'ClientAccessible' context. Please use the 'ClientAccessible' context instead."
        }
        else {
            $errorDescription = "WMI return code: $returnValue"
        }
    }

    return [PSCustomObject]@{
        PSTypeName       = "VSS.CreateResult"
        Success          = $success
        ReturnValue      = $returnValue
        ShadowID         = $shadowId
        VolumePath       = $ResolvedVolume.InputPath
        DeviceID         = $ResolvedVolume.DeviceID
        Context          = $Context
        Description      = ""
        ErrorDescription = $errorDescription
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
    .DESCRIPTION
        Returns VSS.ShadowStorage objects describing the diff area (MaxSpace,
        AllocatedSpace, UsedSpace) associated with each volume. When -VolumePath
        is supplied, only the matching volume is returned. The MaxSpaceGB value
        is rendered as the string "UNLIMITED" when the OS reports the sentinel
        UInt64.MaxValue / 0 value, which is the common state on Windows Client
        editions that have never had their storage limit explicitly set.
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

            $canonicalVolId = ConvertTo-VSSCanonicalVolumeName -Path $volId
            $canonicalDiffVolId = ConvertTo-VSSCanonicalVolumeName -Path $diffVolId

            # Filter if a specific VolumePath was requested
            if ($canonicalDeviceId -and $canonicalVolId -ne $canonicalDeviceId) {
                continue
            }

            # Map to drive letters
            $volInfo = $allvols | Where-Object { (ConvertTo-VSSCanonicalVolumeName -Path $_.DeviceID) -eq $canonicalVolId }
            $diffVolInfo = $allvols | Where-Object { (ConvertTo-VSSCanonicalVolumeName -Path $_.DeviceID) -eq $canonicalDiffVolId }

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
    .DESCRIPTION
        Configures the diff area (shadow storage) for a volume. On Windows Client editions where
        the OS destroys the association when the last snapshot is deleted, this function transparently
        creates a temporary shadow copy to force the association to initialize, applies the limit,
        and removes the temporary snapshot.
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
            if ($_.Volume -match 'DeviceID="(?<id>[^"]+)"') { $volId = $Matches.id }
            (ConvertTo-VSSCanonicalVolumeName -Path $volId) -eq $canonicalVol
        }

        $success = $false
        if ($storage) {
            # Modify existing MaxSpace
            $storage.MaxSpace = $limit
            $result = $storage.Put()
            $success = ($null -ne $result)
        } else {
            # Create new association
            # WMI object path references require backslashes to be doubled (escaped)
            $escapedVol = $resolvedVol.DeviceID.Replace('\', '\\')
            $escapedDiff = $resolvedDiff.DeviceID.Replace('\', '\\')
            $volRef = "Win32_Volume.DeviceID=""$escapedVol"""
            $diffRef = "Win32_Volume.DeviceID=""$escapedDiff"""

            try {
                $class = [wmiclass]"root\cimv2:Win32_ShadowStorage"
                $result = $class.Create($volRef, $diffRef, $limit)
                $returnValue = if ($result -and $result.PSObject.Properties["ReturnValue"]) { [int]$result.ReturnValue } else { [int]$result }
                if ($returnValue -eq 10) {
                    throw "WMI returned 10 (Unknown error)"
                }
                $success = ($returnValue -eq 0)
            } catch {
                # Fallback to vssadmin if WMI fails or complains about method signatures
                $volLetter = $resolvedVol.DriveLetter
                $diffVolLetter = $resolvedDiff.DriveLetter
                $limitStr = if ($MaxSpaceBytes -lt 0) { "UNLIMITED" } else { "$($MaxSpaceBytes)" }
                $cmdResult = Invoke-NativeCommand -CommandPath 'cmd.exe' -Arguments @('/c', 'vssadmin', 'add', 'shadowstorage', "/For=$volLetter", "/On=$diffVolLetter", "/Max=$limitStr")
                $success = ($LASTEXITCODE -eq 0)
                if (-not $success) {
                    # If we fail to add shadow storage directly (common on Windows Client),
                    # attempt the temporary shadow copy workaround to initialize the association.
                    Write-Verbose "Direct shadow storage creation failed. Attempting temporary shadow copy workaround on $($resolvedVol.DeviceID)..."

                    $tempShadow = $null
                    try {
                        # Create a temporary shadow copy to force the OS to initialize the storage association
                        $tempShadow = New-VSSShadowCopy -VolumePath $resolvedVol.DeviceID -Context "ClientAccessible" -Description "Temporary Shadow Copy for Storage Init" -Confirm:$false -ErrorAction Stop

                        if ($tempShadow -and $tempShadow.Success -and $tempShadow.ShadowID) {
                            # Look up the newly created storage association
                            $newStorage = Get-WmiObject -Class Win32_ShadowStorage | Where-Object {
                                $volId = $null
                                if ($_.Volume -match 'DeviceID="(?<id>[^"]+)"') { $volId = $Matches.id }
                                (ConvertTo-VSSCanonicalVolumeName -Path $volId) -eq $canonicalVol
                            }

                            if ($newStorage) {
                                $newStorage.MaxSpace = $limit
                                $putResult = $newStorage.Put()
                                $success = ($null -ne $putResult)
                            }
                        }
                    }
                    catch {
                        Write-Warning "Temporary shadow copy workaround failed: $($_.Exception.Message)"
                    }
                    finally {
                        # Always clean up the temporary shadow copy if it was successfully created
                        if ($tempShadow -and $tempShadow.Success -and $tempShadow.ShadowID) {
                            try {
                                Remove-VSSShadowCopy -VolumePath $resolvedVol.DeviceID -ShadowCopyID $tempShadow.ShadowID -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
                            }
                            catch {
                                Write-Warning "Failed to delete temporary shadow copy $($tempShadow.ShadowID): $($_.Exception.Message)"
                            }
                        }
                    }

                    if (-not $success) {
                        if ($cmdResult -match "Invalid command") {
                             throw "Shadow storage cannot be created manually on Windows Client editions. Windows automatically creates the storage association when the first shadow copy is taken on the volume. Please create a shadow copy on this volume first, then you will be able to change its storage limit."
                        } else {
                             throw "Failed to create shadow storage via WMI, vssadmin, or temporary shadow copy workaround: $cmdResult"
                        }
                    }
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
    .DESCRIPTION
        Tries to parse vssadmin.exe output first (provides rich state info but is locale-
        dependent). If that returns no writers from non-empty output, the parser regex
        didn't match — usually because the OS is non-English — and we fall back to the
        COM-based enumeration via VssBackupHelper.exe, which is locale-independent.
        Pass -ForceCom to skip vssadmin and go straight to COM.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [switch]$ForceCom
    )

    try {
        if (-not (Test-VSSAdministrator)) {
            throw "Administrator privileges are required to query VSS Writers."
        }

        if (-not $ForceCom) {
            $textResult = Get-VSSWritersViaVssAdmin
            if ($null -ne $textResult -and @($textResult).Count -gt 0) {
                return $textResult
            }
            # text parser returned 0 writers — check if the output was non-empty
            # so we know it was a parsing failure (locale) rather than a real empty set.
            if ($script:lastVssAdminOutputWasEmpty) {
                return @()
            }
            Write-Verbose "vssadmin output was non-empty but no writers were parsed; falling back to COM-based enumeration."
        }

        return Get-VSSWritersViaCom
    }
    catch {
        Write-Error "Failed to retrieve VSS Writers: $($_.Exception.Message)"
    }
}

function Get-VSSWritersViaVssAdmin {
    <#
    .SYNOPSIS
        Parses vssadmin list writers output. Internal helper used by Get-VSSWriters.
    .DESCRIPTION
        Captures stdout from `vssadmin list writers` and walks each line, matching
        English-localized "Writer name:", "Writer Id:", "State:" and "Last error:"
        lines to build VSS.Writer objects. The output is locale-dependent; on
        non-English Windows the regex won't match and Get-VSSWriters will fall
        back to Get-VSSWritersViaCom. Sets $script:lastVssAdminOutputWasEmpty
        so the caller can distinguish "no writers reported" from "parsing failed
        because output wasn't in English".
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    $vssadminPath = Join-Path $env:SystemRoot "System32\vssadmin.exe"
    if (-not (Test-Path $vssadminPath)) {
        throw "vssadmin.exe not found on this system. Unable to query VSS Writers."
    }

    $output = Invoke-NativeCommand -CommandPath vssadmin -Arguments @('list', 'writers')
    $script:lastVssAdminOutputWasEmpty = ($output.Count -eq 0 -or ($output -join "`n").Trim().Length -eq 0)
    if ($script:lastVssAdminOutputWasEmpty) {
        return @()
    }

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
            if ($current.StateCode -eq 1) {
                $current.StatusColor = "#4CAF50"
            } else {
                $current.StatusColor = "#FF9800"
            }
        }
        elseif ($trimmed -match "Last error:\s*(?<Error>.*)") {
            $current.LastError = $Matches.Error.Trim()
            if ($current.LastError -ne "No error") {
                $current.StatusColor = "#F44336"
            }
        }
    }

    if ($null -ne $current) {
        $list += [PSCustomObject]$current
    }

    $list
}

function Get-VSSWritersViaCom {
    <#
    .SYNOPSIS
        Queries VSS Writers via the COM API. Used as a fallback when vssadmin output parsing
        fails (e.g. on non-English Windows where the localized output doesn't match the parser regex).
    .DESCRIPTION
        Spawns VssBackupHelper.exe with the --list-writers argument, which gathers writer metadata
        and returns the result as JSON on stdout. Requires VssBackupHelper.exe to be present in
        bin\ or to be compileable.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    $helperPath = Get-VSSBackupHelperPath
    if (-not $helperPath) {
        throw "VssBackupHelper.exe is required for COM-based writer enumeration but is not available."
    }

    $output = Invoke-NativeCommand -CommandPath $helperPath -Arguments @('--list-writers')
    if ($LASTEXITCODE -ne 0) {
        throw "VssBackupHelper.exe returned exit code $LASTEXITCODE. Output: $($output -join "`n")"
    }

    $json = $output -join "`n"
    $payload = $json | ConvertFrom-Json -ErrorAction Stop
    foreach ($writer in $payload) {
        $statusColor = if ($writer.state -eq "Stable" -or $writer.stateCode -eq 1) { "#4CAF50" } else { "#FF9800" }
        # lastError is only present in vssadmin-derived output; the COM JSON doesn't
        # include it, so we read it conditionally.
        if ($writer.PSObject.Properties['lastError'] -and $writer.lastError -and $writer.lastError -ne "No error") {
            $statusColor = "#F44336"
        }
        $lastError = if ($writer.PSObject.Properties['lastError']) { [string]$writer.lastError } else { "N/A (COM)" }
        [PSCustomObject]@{
            PSTypeName  = "VSS.Writer"
            WriterName  = $writer.name
            WriterId    = $writer.id
            State       = $writer.state
            StateCode   = $writer.stateCode
            LastError   = $lastError
            StatusColor = $statusColor
        }
    }
}

function Mount-VSSShadowCopy {
    <#
    .SYNOPSIS
        Mounts a VSS shadow copy as a folder symlink.
    .DESCRIPTION
        Creates a symbolic link at the specified MountPath that points to the shadow copy's
        DeviceObject. The shadow copy's contents become accessible through that folder. Use
        Dismount-VSSShadowCopy to remove the symlink (the underlying shadow copy is not
        affected; delete it with Remove-VSSShadowCopy to free the space).
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

        $shadow = Get-WmiObject -Class Win32_ShadowCopy -Filter "ID='$ShadowCopyID'"
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
            Invoke-NativeCommand -CommandPath 'cmd.exe' -Arguments @('/c', 'rmdir', "$MountPath") | Out-Null
            Remove-Item -Path $MountPath -Force -Recurse -ErrorAction SilentlyContinue | Out-Null
        }

        # Use cmd mklink /d since PowerShell's New-Item validates target paths, which fails for raw VSS devices
        $cmdResult = Invoke-NativeCommand -CommandPath 'cmd.exe' -Arguments @('/c', 'mklink', '/d', "$MountPath", "$target")
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
    .DESCRIPTION
        Removes only the symbolic link at the given MountPath; the underlying shadow copy is
        unaffected (use Remove-VSSShadowCopy to delete the snapshot itself).
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
            $result = Invoke-NativeCommand -CommandPath 'cmd.exe' -Arguments @('/c', 'rmdir', "$MountPath")
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

function Export-VSSReport {
    <#
    .SYNOPSIS
        Exports volumes, shadow copies, and/or VSS writers to a CSV or JSON file.
    .DESCRIPTION
        Useful for sysadmins who need a snapshot of the VSS state for audit, reporting,
        or trend analysis. -Kind selects what to export; -Path is the output file; -Format
        defaults to CSV if -Path ends in .csv, JSON if it ends in .json, otherwise CSV.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Volumes', 'ShadowCopies', 'Writers')]
        [string]$Kind,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [string]$VolumePath,

        [ValidateSet('CSV', 'JSON')]
        [string]$Format
    )

    if (-not $Format) {
        if ($Path -like '*.json') { $Format = 'JSON' } else { $Format = 'CSV' }
    }

    switch ($Kind) {
        'Volumes'     { $data = @(Get-VSSSupportedVolumes) }
        'ShadowCopies' {
            if (-not $VolumePath) {
                throw "Export-VSSReport -Kind ShadowCopies requires -VolumePath."
            }
            $data = @(Get-VSSShadowCopies -VolumePath $VolumePath)
        }
        'Writers'     { $data = @(Get-VSSWriters) }
    }

    if (-not $data -or $data.Count -eq 0) {
        Write-Warning "No data to export; writing an empty $Format file at $Path."
    }

    $parent = Split-Path -Path $Path -Parent
    if ($parent -and -not (Test-Path -Path $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    switch ($Format) {
        'CSV'  { $data | Export-Csv -Path $Path -NoTypeInformation -Force }
        'JSON' { $data | ConvertTo-Json -Depth 5 | Out-File -FilePath $Path -Encoding UTF8 -Force }
    }

    return $data.Count
}

# Explicit module export. FunctionsToExport in the .psd1 manifest is the authoritative
# list for external callers; Export-ModuleMember here reinforces the boundary between
# the public API (everything above) and any future private helpers we add below this line.
Export-ModuleMember -Function @(
    'Test-VSSAdministrator'
    'ConvertTo-VSSCanonicalVolumeName'
    'Get-VSSUsageColor'
    'Get-VSSShadowStateName'
    'Get-VSSShadowStateColor'
    'ConvertFrom-VSSWmiDateTime'
    'Get-VSSAgeText'
    'Resolve-VSSVolumePath'
    'Get-VSSSupportedVolumes'
    'Get-VSSShadowCopies'
    'Get-VSSBackupHelperPath'
    'New-VSSShadowCopy'
    'Remove-VSSShadowCopy'
    'Get-VSSShadowStorage'
    'Set-VSSShadowStorageLimit'
    'Get-VSSWriters'
    'Get-VSSWritersViaCom'
    'Get-VSSWritersViaVssAdmin'
    'Mount-VSSShadowCopy'
    'Dismount-VSSShadowCopy'
    'Export-VSSReport'
) -Variable @('ModuleRoot')

# Note: Get-VSSCreationStrategy, New-VSSShadowCopyViaHelper, New-VSSShadowCopyViaDiskShadow,
# and New-VSSShadowCopyViaWmi are deliberately not exported — they exist to let each
# shadow-copy creation strategy be unit-tested in isolation.
