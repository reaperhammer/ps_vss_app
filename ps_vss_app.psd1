@{
    RootModule        = 'ps_vss_app.psm1'
    ModuleVersion     = '1.1.0'
    GUID              = '2b5d587d-9917-4d35-b830-8487d2ee6fe2'
    Author            = 'ps_vss_app contributors'
    Copyright         = '(c) ps_vss_app contributors. Licensed under the Apache 2.0 License.'
    Description       = 'Volume Shadow Copy Service (VSS) management for Windows. Provides WMI-based volume/shadow copy queries, a dynamically compiled C# COM Interop helper for client OS writer-consistent snapshots, shadow storage configuration, and a WPF management GUI.'
    PowerShellVersion = '5.1'
    # The module is compatible with both Windows PowerShell 5.1 (Desktop) and
    # PowerShell 7+ (Core). The WPF GUI in VSSManager.ps1 is restricted to
    # 5.1 via an in-script #Requires check because System.Windows is not
    # available in Core editions.
    CompatiblePSEditions = @('Desktop', 'Core')
    CompanyName       = 'ps_vss_app contributors'

    # Functions exported by this module. The GUI (VSSManager.ps1) and external
    # callers can use these directly after Import-Module ps_vss_app.
    FunctionsToExport = @(
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
    )

    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    # Required assemblies: WMI for volume queries, the .NET 4.0 compiler for
    # dynamic compilation of the C# helper when bin\VssBackupHelper.exe is
    # missing.
    RequiredAssemblies = @(
        'System.Management'
    )

    PrivateData = @{
        PSData = @{
            Tags       = @('VSS', 'ShadowCopy', 'Backup', 'Windows', 'WMI', 'WPF', 'Snapshot', 'VolumeShadowCopy')
            ProjectUri = 'https://github.com/ps_vss_app/ps_vss_app'
            LicenseUri = 'https://github.com/ps_vss_app/ps_vss_app/blob/main/LICENSE'
        }
    }
}
