# VSS Manager - Issues & Resolutions List

All issues identified during testing on Windows Client have been investigated and resolved:

## 1. VSS Shadow Storage Limits on Windows Client
- **Observed Behavior**:
  - Setting a limit succeeds when a snapshot exists on the drive, but fails if all snapshots are deleted (the storage association disappears when all snapshots are deleted).
  - Even after successfully saving a limit (e.g. 10GB), returning to the Volume Management tab details section still reports the limit as "unlimited".
- **Resolutions**:
  - [x] **Investigate if the "unlimited" display is a GUI query/refresh bug**: Resolved. WMI returns volume reference properties (like `Volume` and `DiffVolume` in `Win32_ShadowStorage`) with escaped backslashes (e.g., `\\\\?\\Volume{...}\\`), resulting in 4 backslashes in parsing. The query filter compared this directly to canonical paths (which use 2 backslashes), causing the query to return empty results and the GUI to default to displaying "unlimited". We resolved this by canonicalizing both paths via `ConvertTo-VSSCanonicalVolumeName` before comparison and drive letter lookup.
  - [x] **Confirm if the shadow storage association is completely destroyed by Windows and design a workaround**: Resolved. On Windows Client editions, if a custom limit was never explicitly set, Windows automatically destroys the temporary shadow storage association when the last snapshot is deleted. Trying to set a limit when no association exists fails on Client because `vssadmin add shadowstorage` is not supported. Workaround: If no association exists when setting a limit, the script now temporarily creates a shadow copy to force the OS to initialize the association, applies the custom limit, and then deletes the temporary shadow copy. Since a custom limit is now set, the association becomes persistent and is retained by the OS.

## 2. Backup Context Creation Failure
- **Observed Behavior**:
  - Creating a shadow copy using the `Backup` context fails on Windows Client.
- **Resolutions**:
  - [x] **Investigate the exact error code/reason for failure under the `Backup` context**: Resolved. The WMI `Win32_ShadowCopy.Create` method returns error code `5` (Unsupported shadow copy context).
  - [x] **Check if this is due to specific limitations of Windows Client OS**: Resolved. The `Backup` context is not supported by the WMI `Win32_ShadowCopy` class on Windows Client editions (which only support the `ClientAccessible` context).
  - [x] **Implement dynamic C# COM Interop compilation solution for Client OS**: Resolved. Built a pure C# command line requester utility `VssBackupHelper.cs` (compiled dynamically to `bin\VssBackupHelper.exe`) which directly imports the VSS COM interface (`IVssBackupComponents`) from `vssapi.dll`.
  - [x] **Verify writer-consistent/application-consistent backups on Windows Client**: Resolved. The compiled tool successfully freezes system writers and creates shadow copies under both `Backup` (0) and `ClientAccessibleWriters` (13) contexts on Windows Client OS without requiring any server-only utilities (`diskshadow.exe`) or external libraries (AlphaVSS).
  - [x] **Re-enable Backup context in the GUI**: Resolved. Updated `VSSManager.ps1` to re-enable the "Backup" context selection option in the combobox now that it is fully supported via `VssBackupHelper`.
