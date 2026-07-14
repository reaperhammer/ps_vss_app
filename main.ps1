# main.ps1 - Backward-compatibility shim for ps_vss_app.
#
# This file used to host the VSS helper functions directly. As of v1.1.0 the canonical
# home of those functions is the ps_vss_app.psm1 module (loaded by ps_vss_app.psd1).
#
# This shim is preserved so existing scripts and READMEs that document
# `. .\main.ps1` continue to work; `Import-Module -Global` exposes all of the
# module's public functions in the global scope, which is what dot-sourcing did
# historically.

#Requires -Version 5.1

[CmdletBinding()]
param()

$manifestPath = Join-Path $PSScriptRoot 'ps_vss_app.psd1'
if (-not (Test-Path $manifestPath)) {
    throw "ps_vss_app.psd1 not found at $manifestPath. The module must be installed alongside main.ps1."
}

Import-Module $manifestPath -Force -Global
