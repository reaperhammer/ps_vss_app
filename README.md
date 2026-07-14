# VSS Manager - Volume Shadow Copy Service Manager

A modern PowerShell-based application for managing Windows Volume Shadow Copy Service (VSS) with both command-line and graphical user interfaces.

## Features

### 🖥️ Modern WPF GUI
- **Volume Management Tab**: View all VSS-supported volumes with detailed statistics
- **Shadow Copy Operations Tab**: Create, view, and delete shadow copies
- **Real-time Progress Tracking**: Visual progress indicators for all operations
- **Usage Statistics**: Color-coded volume capacity bars and usage percentages

### 📋 Core Functionality
- **List VSS-Supported Volumes**: Display all volumes that support shadow copies
- **View Shadow Copies**: Show existing shadow copies with creation time, state, and age
- **Create Shadow Copies**: Generate new volume snapshots with progress tracking  
- **Delete Shadow Copies**: Remove individual or all shadow copies for a volume
- **Volume Statistics**: Display capacity, free space, usage percentages, and file systems
- **VSS Context Flexibility**: Support for `ClientAccessible`, `Backup`, `ClientAccessibleWriters`, and `AppRollback` contexts

### ⚙️ Client-Side VSS Writer Support (Dynamic COM Interop)
- Windows Client editions (e.g. Windows 10/11) do not support the server-only `diskshadow.exe` utility and limit WMI-based shadow copy creation to the `ClientAccessible` context (which ignores VSS writers).
- VSS Manager resolves this limitation natively using a dynamically compiled C# helper utility ([VssBackupHelper.cs](VssBackupHelper.cs)).
- The utility directly consumes the unmanaged VSS COM API (`IVssBackupComponents` from `vssapi.dll`), allowing full writer-freeze, application-consistent backups on client editions of Windows without requiring any external libraries or VC++ redistributables.
- The wrapper is compiled on the fly at startup if `bin\VssBackupHelper.exe` is missing.

## Requirements

- **Windows 10/11** (or Windows Server equivalent)
- **Windows PowerShell 5.1** (not compatible with PowerShell 7/Core)
- **Administrator privileges** (required for VSS operations)
- **.NET Framework** (for WPF GUI and dynamic compilation components)

## Installation & Usage

### Quick Start

1. **Clone or download** this repository
2. **Run**: Double click `Launch.bat`
3. The application will automatically launch the GUI interface

### Command Line Usage

You can also use the core functions directly in PowerShell:

```powershell
# Import the functions
. .\main.ps1

# List all VSS-supported volumes
Get-VSSSupportedVolumes | Format-Table

# View shadow copies for C: drive
Get-VSSShadowCopies -VolumePath "C:" | Format-Table

# Create a new shadow copy with default context
New-VSSShadowCopy -VolumePath "C:"

# Create an application-consistent backup shadow copy with VSS writers active
New-VSSShadowCopy -VolumePath "C:" -Context "Backup"

# Preview shadow copy deletion without making changes
Remove-VSSShadowCopy -VolumePath "C:" -WhatIf

# Remove shadow copies (with confirmation)
Remove-VSSShadowCopy -VolumePath "C:"
```

### Tests

Run the helper-function test suite with Windows PowerShell:

```powershell
Import-Module Pester
Invoke-Pester -Script .\tests
```

### GUI Interface

The GUI provides two main tabs:

#### Volume Management
- View all available volumes with detailed information
- Color-coded usage indicators (Green < 70%, Orange < 85%, Red ≥ 85%)
- Select volumes for shadow copy operations

#### Shadow Copy Operations  
- View existing shadow copies for selected volumes
- Create new shadow copies with progress tracking (fully supporting `Backup` and `ClientAccessible` contexts)
- Delete selected shadow copies with confirmation dialogs
- Real-time status updates and operation feedback

## File Structure

```
ps_vss_app/
├── main.ps1           # Core VSS functions and compilation routines
├── VSSManager.ps1     # WPF GUI application
├── VssBackupHelper.cs # Native C# VSS COM Interop requester source
├── Launch.bat         # Batch file launcher
├── LICENSE           # Apache 2.0 License
├── assets/png/       # GUI icons and images
└── bin/              # Location of compiled VssBackupHelper.exe
```

## Compatibility Notes

- **PowerShell Version**: This application requires Windows PowerShell 5.1 and will not work with PowerShell 7 (Core)
- **Platform**: Windows-only (uses Windows Management Instrumentation and .NET/COM)
- **Architecture**: Compatible with both x64 and x86 Windows systems

## Troubleshooting

### Common Issues

1. **"Script cannot be loaded" error**
   - Run: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

2. **"Access Denied" errors**
   - Ensure you're running as Administrator
   - Check that VSS service is running: `sc query vss`

3. **GUI doesn't appear**
   - Verify you're using Windows PowerShell 5.1 (not PowerShell 7)
   - Check that .NET Framework is installed

4. **Missing PNG files**
   - The application will create the assets directory automatically
   - The GUI will start even if some PNG assets are missing; missing toolbar/tab icons render as blank cells, and missing image sources are cleared automatically
   - `assets/png/app_icon.png` is the only icon that previously caused a hard failure during XAML parsing; this is now handled safely and the window simply uses the default icon if it is missing

## Security & Permissions

This application requires administrative privileges because:
- Volume Shadow Copy operations require elevated permissions
- WMI queries for system volume information need admin access
- Shadow copy creation/deletion affects system-level services

## License

Licensed under the Apache License, Version 2.0. See [LICENSE](LICENSE) file for details.

## Contributing

This is a system administration tool for Windows VSS management. Contributions should focus on:
- Improving error handling and user experience
- Adding additional VSS-related functionality
- Enhancing the GUI interface
- Bug fixes and performance improvements

---

**⚠️ Important**: Always test shadow copy operations in a non-production environment first. Volume Shadow Copy operations can affect system performance and disk space.