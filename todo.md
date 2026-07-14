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

---

## Improvement Backlog

Forward-looking enhancements identified during review, roughly ordered by impact. Items marked ⭐ are the recommended starting point.

### High impact

- [x] ⭐ **Add a PowerShell module manifest** (`ps_vss_app.psd1` + `.psm1`). Refactor `main.ps1` into a proper module. The GUI's `. $mainScript` dot-source should become `Import-Module`. Enables `Get-Command -Module ps_vss_app` discoverability, versioning, and proper `Export-ModuleMember` boundaries.
- [x] ⭐ **Modernize the Pester test suite for Pester 5+**. `tests/main.Tests.ps1` uses Pester 3 syntax (`Should Be`, blanket `function vssadmin {}` placeholders). Won't run on a default Pester 5 install. Update to `Should -Be`, use `InModuleScope` mocks, and add a `Build.ps1` wrapper.
- [x] ⭐ **Add CI** (`.github/workflows/test.yml`). Run `Invoke-Pester ./tests` + `PSScriptAnalyzer` on every PR. Catches the Pester 5 issue above immediately and would have flagged several of the items below earlier.
- [x] **Split `New-VSSShadowCopy` into smaller strategies**. The single function is ~150 lines handling helper / diskshadow / WMI paths with three different error decoders. Extract `Get-VSSCreationStrategy`, `New-VSSShadowCopyViaHelper`, `New-VSSShadowCopyViaDiskShadow`, `New-VSSShadowCopyViaWmi` for testability.
- [x] **Decouple the C# helper compiler path from `csc.exe`**. `Get-VSSBackupHelperPath` hard-codes `Framework64\v4.0.30319\csc.exe`. On machines without the .NET Framework SDK it silently returns `$null`, then users get a confusing "context 5" WMI error on Client. Add Roslyn (`Microsoft.CodeAnalysis.CSharp`) as a fallback, and surface a clear error when the helper was needed but couldn't be built.

### Code quality

- [x] **Replace text-based `Get-VSSWriters` with a COM call**. The current regex parser for `vssadmin list writers` stdout breaks on any non-English Windows. Use `GatherWriterMetadata` + `GetWriterMetadataCount` + `GetWriterMetadata` from the existing C# helper — it already imports the interface.
- [x] **Add missing `.DESCRIPTION` blocks** to `Mount-VSSShadowCopy`, `Dismount-VSSShadowCopy`, `Set-VSSShadowStorageLimit`, `Get-VSSWriters`. `Get-Help` is half-broken for these.
- [x] **Unify `Resolve-VSSVolumePath` return shape** — drop the `.Volume` field that exposes the raw WMI object, or stop returning the PSCustomObject. The current dual shape forces Pester tests to mock the whole return value.
- [x] **Fix `Launch.bat` CWD** — add `cd /d "%~dp0"` before `Start-Process`. The XAML uses relative paths (`.\assets\png\app_icon.png`) that resolve against the *process* CWD, not the script dir, so launching from elsewhere causes WPF parse errors.
- [x] **Add `#Requires -Version 5.1`** at the top of `main.ps1` and `VSSManager.ps1` so non-5.1 hosts fail with a clear message instead of WMI warnings.

### UX / UI

- [x] **Keyboard shortcuts**: `F5` refresh, `Ctrl+N` new shadow copy, `Del` delete selected.
- [x] **Enable the Browse button** for mounted shadow copies (currently `IsEnabled="False"` with no tooltip explaining why) and add an "Open in Explorer" action using the existing `explorer.png` icon asset.
- [x] **CSV / JSON export** of the volumes, shadow copies, and writers lists — sysadmins always want this for reports.
- [x] **Show real progress** in the status bar. `VssBackupHelper` has discrete stages (GatherMetadata / Prepare / DoSnapshot) that map to real percentages, but the GUI uses an indeterminate bar.

### Nice-to-haves

- [x] **`CHANGELOG.md`** — `git log` shows real features (advanced VSS, the C# requester, the storage workaround) that a release-aware changelog would make easier to track.
- [x] **Tagged release artifacts** — publish the compiled `VssBackupHelper.exe` as a GitHub release asset for users who want to skip the on-first-run compile.
- [x] **Switch `ConvertTo-VSSCanonicalVolumeName` to `StringComparer.OrdinalIgnoreCase`** instead of uppercasing both sides — slightly faster, slightly more idiomatic.
- [x] **Dark mode** for the WPF UI (via `Window.Background` style swap or a system-follow toggle).
