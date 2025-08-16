# PowerShell script for Volume Shadow Copy (VSS) operations on Windows 11 using WMI
# Date: Wednesday, August 13, 2025
# Time: 17:46:19

$DebugPreference = "SilentlyContinue"

# Function 1: List volumes that support VSS
function Get-VSSSupportedVolumes {
    <#
    .SYNOPSIS
        Lists volumes that support Volume Shadow Copy Service (VSS)
    .DESCRIPTION
        Lists all volumes that support VSS, showing their properties and indicating VSS support status
    #>
    try {
        $volumes = Get-WmiObject -Class Win32_Volume -ErrorAction Stop
        
        Write-Host "Supported Volumes:" -ForegroundColor Green
        Write-Host "==================" -ForegroundColor Green
        
        foreach ($volume in $volumes) {
            if ($volume.DriveLetter -and $volume.DriveType -eq 3) {
                Write-Host "Drive Letter: $($volume.DriveLetter)" -ForegroundColor Yellow
                Write-Host "  Device ID: $($volume.DeviceID)" 
                Write-Host "  Volume Name: $($volume.Name)"
                Write-Host "  File System: $($volume.FileSystem)"
                Write-Host "  Capacity: $([math]::Round($volume.Capacity / 1GB, 2)) GB"
                Write-Host "  Free Space: $([math]::Round($volume.FreeSpace / 1GB, 2)) GB"
                Write-Host "  Label: $($volume.Label)"
                Write-Host ""
            }
        }
    }
    catch {
        Write-Error "Failed to retrieve volumes: $($_.Exception.Message)"
    }
}


# Function 2: List shadow copies for a specific volume
function Get-VSSShadowCopies {
    param(
        [Parameter(Mandatory=$true)]
        [string]$VolumePath
    )
    <#
    .SYNOPSIS
        Lists all shadow copies for a specific volume
    .DESCRIPTION
        Retrieves and displays information about all shadow copies associated with the specified volume
    #>
    try {
        # Normalize volume path format
        $normalizedPath = $VolumePath.replace("\","").replace(":","").Trim()
        if ($normalizedPath -match '^[A-Za-z]$') {
        Write-Debug "Normal Drive"
            # It's a drive letter, convert to device ID format for comparison
            $driveLetter = $normalizedPath + ":"
            $volumes = Get-WmiObject -Class Win32_Volume -ErrorAction Stop | Where-Object { $_.DriveLetter -eq $driveLetter }
            
            if ($volumes) {
                $deviceID = $volumes.DeviceID
                $shadowCopies = Get-WmiObject -Class Win32_ShadowCopy -ErrorAction Stop | Where-Object { $_.VolumeName -eq "$deviceID" }
                Write-Debug "$($shadowCopies.count) Shadow Found"
            } else {
                Write-Warning "No volume found for drive letter: $driveLetter"
                return
            }
        } else {
        Write-Debug "DeviceID Drive"
            # Assume it's a device ID or full path
            $shadowCopies = Get-WmiObject -Class Win32_ShadowCopy -ErrorAction Stop | Where-Object { $_.VolumeName -eq "$VolumePath" }
        }
        
        if ($shadowCopies) {
            Write-Host "Found $($shadowCopies.count) Shadow Copies for Volume: $VolumePath" -ForegroundColor Green
            Write-Host "==================================" -ForegroundColor Green
            
            foreach ($shadow in $shadowCopies) {
                Write-Host "Shadow Copy ID: $($shadow.ID)" -ForegroundColor Yellow
                Write-Host "  Device Object: $($shadow.DeviceObject)"
                Write-Host "  Volume Name: $($shadow.VolumeName)"
                Write-Host "  Creation Time: $($shadow.ConvertToDateTime($shadow.InstallDate).ToLocalTime().Tostring())" 
                Write-Host "  State: $($shadow.State)"
                Write-Host "  Persistent: $($shadow.Persistent)"
                Write-Host "  Client Accessible: $($shadow.ClientAccessible)"
                Write-Host ""
            }
        } else {
            Write-Host "No shadow copies found for volume: $VolumePath" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "Failed to list shadow copies for volume `$VolumePath`: `$_"
    }
}


# Function 3: Create a new shadow copy
function New-VSSShadowCopy {
    param(
        [Parameter(Mandatory=$true)]
        [string]$VolumePath,
        [string]$Description = "PowerShell Created Shadow Copy"
    )
    
    <#
    .SYNOPSIS
        Creates a new shadow copy for a specified volume
    .DESCRIPTION
        Creates a new Volume Shadow Copy for the specified volume with optional description
    #>
    try {

    #################################

        # Normalize volume path format
        $normalizedPath = $VolumePath.replace("\","").replace(":","").Trim()
        if ($normalizedPath -match '^[A-Za-z]$') {
            # It's a drive letter, convert to device ID format
            $driveLetter = $normalizedPath + ":"
            $volumes = Get-WmiObject -Class Win32_Volume -ErrorAction Stop | Where-Object { $_.DriveLetter -eq $driveLetter }
            
            if ($volumes) {
                $deviceID = $volumes.DeviceID
            } else {
                Write-Error "No volume found for drive letter: $driveLetter"
                return
            }
        } else {
            # Assume it's a device ID or full path
            $deviceID = $VolumePath
        }
        
        # Use Win32_ShadowCopy.Create method (requires proper COM registration)
        Write-Host "Creating shadow copy for Drive $normalizedPath volume: $deviceID" -ForegroundColor Yellow

        $shadowpath = (Get-WmiObject -list win32_shadowcopy).Create("$($deviceID)","ClientAccessible") | select-object ShadowID -ExpandProperty ShadowID

        if ($shadowpath) {
        Write-Host "Successfully created shadow copy $shadowpath for Drive $normalizedPath volume: $deviceID" -ForegroundColor Yellow
        } else {
        Write-Warning "Shadow copy creation requires administrative privileges and proper VSS setup."
        }

    }
    catch {
        Write-Error "Failed to create shadow copy for volume $VolumePath`: $_"
    }
}


# Function 4: Remove shadow copies
function Remove-VSSShadowCopy {
    param(
        [Parameter(Mandatory=$true)]
        [string]$VolumePath,
        [string]$ShadowCopyID = $null,
        [switch]$Confirm = $true
    )
    
    <#
    .SYNOPSIS
        Removes shadow copies for a specific volume
    .DESCRIPTION
        Deletes either a specific shadow copy or all shadow copies for a specified volume
    #>

    try {
        # Get shadow copies matching the criteria
        if ($ShadowCopyID) {
            $shadowCopies = Get-WmiObject -Class Win32_ShadowCopy -ErrorAction Stop | Where-Object { $_.ID -eq $ShadowCopyID }
        } else {
            # Normalize volume path format
            $normalizedPath = $VolumePath.replace("\","").replace(":","").Trim()
            if ($normalizedPath -match '^[A-Za-z]$') {
                $driveLetter = $normalizedPath + ":"
                $volumes = Get-WmiObject -Class Win32_Volume -ErrorAction Stop | Where-Object { $_.DriveLetter -eq $driveLetter }
                
                if ($volumes) {
                    $deviceID = $volumes.DeviceID
                    $shadowCopies = Get-WmiObject -Class Win32_ShadowCopy -ErrorAction Stop | Where-Object { $_.VolumeName -like "*$deviceID*" }
                } else {
                    Write-Warning "No volume found for drive letter: $driveLetter"
                    return
                }
            } else {
                $shadowCopies = Get-WmiObject -Class Win32_ShadowCopy -ErrorAction Stop | Where-Object { $_.VolumeName -like "*$VolumePath*" }
            }
        }
        
        if ($shadowCopies) {
            if ($Confirm) {
                Write-Host "The following shadow copies will be deleted:" -ForegroundColor Red
                foreach ($shadow in $shadowCopies) {
                    Write-Host "  ID: $($shadow.ID) - Created: $($shadow.ConvertToDateTime($shadow.InstallDate).ToLocalTime().Tostring())" -ForegroundColor Yellow
                }
                
                $confirmation = Read-Host "`nAre you sure you want to proceed with deletion? (Y/N)"
                if ($confirmation -ne 'Y' -and $confirmation -ne 'y') {
                    Write-Host "Deletion cancelled." -ForegroundColor Yellow
                    return
                }
            }
            
            # Delete shadow copies
            foreach ($shadow in $shadowCopies) {
                try {
                    Write-Host "Deleting Shadow Copy ID: $($shadow.ID)" -ForegroundColor Red
                    
                    # Use the Delete method of Win32_ShadowCopy
                    $result = $shadow.Delete()
                    
                    if (!($result)) {
                        Write-Host "Shadow Copy $($shadow.ID) deleted successfully!" -ForegroundColor Green
                    } else {
                        Write-Warning "Failed to delete shadow copy $($shadow.ID). Return code: $($result.ReturnValue)"
                    }
                } catch {
                    Write-Error "Failed to delete shadow copy $($shadow.ID): $_"
                }
            }
        } else {
            Write-Host "No matching shadow copies found for volume: $VolumePath" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "Failed to delete shadow copies for volume $VolumePath`: $_"
    }
}

# Example usage:
# Get-VSSSupportedVolumes
# Get-VSSShadowCopies -VolumePath "E:\"
# New-VSSShadowCopy -VolumePath "C:\" -Description "Test Shadow Copy"
# Remove-VSSShadowCopy -VolumePath "E:\" -Confirm:$false