# PURPOSE
# -------
# Imports the locally-built Visio module directly from bin/Debug into the current
# PowerShell session. This is the fastest way to test the module without installing
# it -- no copy to the user's PowerShell modules folder.
#
# Run after building the solution in Debug.

Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"

$visio_psd1 = Join-Path $PSScriptRoot ".\bin\Debug\Visio.psd1"
Import-Module $visio_psd1
