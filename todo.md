# VSS Manager - Pending Issues & TODO List

The following issues were identified during testing on Windows Client and need to be resolved in a future session:

## 1. VSS Shadow Storage Limits on Windows Client
- **Observed Behavior**:
  - Setting a limit succeeds when a snapshot exists on the drive, but fails if all snapshots are deleted (the storage association disappears when all snapshots are deleted).
  - Even after successfully saving a limit (e.g. 10GB), returning to the Volume Management tab details section for the G: drive still reports the limit as "unlimited".
- **Tasks**:
  - [ ] Investigate if the "unlimited" display is a GUI query/refresh bug (e.g. in `Get-VSSShadowStorage` matching/parsing) or if WMI's `$storage.Put()` is not committing the value correctly.
  - [ ] Confirm if the shadow storage association is completely destroyed by Windows when the last snapshot is deleted, and design a workaround (e.g. automatically handling clean recreation or retaining the setting).

## 2. Backup Context Creation Failure
- **Observed Behavior**:
  - Creating a shadow copy using the `Backup` context on the G: drive fails.
- **Tasks**:
  - [ ] Investigate the exact error code/reason for failure under the `Backup` context.
  - [ ] Check if this is due to VSS writers vetoing the request, missing administrative options, or specific limitations of Windows Client OS when freezing non-system drives under `Backup` context.
