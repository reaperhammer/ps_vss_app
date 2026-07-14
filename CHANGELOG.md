# Changelog

All notable changes to ps_vss_app are documented in this file. Versions follow
[Semantic Versioning](https://semver.org/spec/v2.0.0.html). Dates are in
YYYY-MM-DD format and reflect the commit date, not the release date.

## [Unreleased]

## [1.1.0] - 2026-07-14

### Added
- PowerShell module manifest (`ps_vss_app.psd1`) and module file (`ps_vss_app.psm1`).
  `main.ps1` is now a backward-compatibility shim that imports the module.
  `Get-Command -Module ps_vss_app` now discovers the public surface.
- Pester 5+ test suite. All 31 tests pass under Pester 5/6 in PowerShell 7.
  `Build.ps1` is a thin wrapper that installs Pester and runs the suite.
- CI workflow at `.github/workflows/test.yml` that runs the Pester suite on
  Windows against PowerShell 7.
- COM-based VSS writer enumeration (`Get-VSSWritersViaCom`) and the
  `--list-writers` mode in `VssBackupHelper.exe`. `Get-VSSWriters` now
  auto-falls-back to the COM path when the vssadmin output can't be parsed
  (typical on non-English Windows).
- `-ForceCom` switch on `Get-VSSWriters` to skip vssadmin and go straight to COM.
- `Export-VSSReport` cmdlet for CSV/JSON export of volumes, shadow copies, and
  writers.
- Keyboard shortcuts in the WPF GUI: F5 refresh, Ctrl+N new shadow copy,
  Del delete selected, Ctrl+E export.
- Export button in the Shadow Copies tab.
- Progress markers (`PROGRESS:<pct>:<message>`) emitted by the C# helper and
  surfaced via `Write-Verbose` from the PowerShell side, so a host can render
  a determinate progress bar during `New-VSSShadowCopy` with the `Backup`
  context.
- Roslyn (`Microsoft.CSharp.CSharpCodeProvider`) fallback in
  `Get-VSSBackupHelperPath` so `VssBackupHelper.exe` can be compiled on hosts
  that don't have `csc.exe` (no .NET Framework SDK installed).
- `Invoke-NativeCommand` private helper that wraps `&` so Pester 5+ can mock
  external command invocations cleanly.

### Changed
- `New-VSSShadowCopy` is now a thin orchestrator that picks one of
  `New-VSSShadowCopyViaHelper`, `New-VSSShadowCopyViaDiskShadow`, or
  `New-VSSShadowCopyViaWmi` based on a `Get-VSSCreationStrategy` decision.
  Each strategy is independently testable.
- `Resolve-VSSVolumePath` exposes the raw `Win32_Volume` WMI object as
  `OriginalWmiObject` (not `Volume` or `WmiObject`), so the public shape is
  explicit about whether you're holding the canonical VSS view or a leaked WMI
  reference.
- Module manifest declares `CompatiblePSEditions = @('Desktop', 'Core')` so
  tests can run under PowerShell 7; the WPF GUI in `VSSManager.ps1` keeps its
  in-script `#Requires -Version 5.1` because System.Windows isn't available in
  Core editions.
- `Launch.bat` now does `cd /d "%~dp0"` before launching PowerShell, so the
  GUI's relative asset paths (e.g. `.\assets\png\app_icon.png`) resolve
  correctly regardless of the caller's CWD.

### Fixed
- `csc.exe` hard-coded path in `Get-VSSBackupHelperPath` no longer silently
  returns `$null` on hosts without the .NET Framework SDK; the Roslyn
  fallback handles those cases and a clear error is surfaced when compilation
  fails completely.
- The `$cscPath:` parser ambiguity in `Get-VSSBackupHelperPath` (PowerShell
  interpreting `$cscPath:` as a drive reference) — fixed with `${cscPath}`.
- Module functions no longer rely on `Set-StrictMode -Version 2.0` to fail
  in subtle ways when invoked from the InModuleScope test scope.
- Set-StrictMode is now applied in the .psm1 (module scope) rather than at
  script-scope, so callers don't inherit the strict mode.

## [1.0.0] - 2026-01-15

Initial release. Pre-dates this changelog; see `git log` for the full history.

[Unreleased]: https://github.com/ps_vss_app/ps_vss_app/compare/v1.1.0...HEAD
[1.1.0]: https://github.com/ps_vss_app/ps_vss_app/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/ps_vss_app/ps_vss_app/releases/tag/v1.0.0
