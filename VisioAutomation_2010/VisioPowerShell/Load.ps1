# Loads the module fom bin debug.
# For those cases where you need to check something quick
# and don't want to overwrite the installed version

Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"

$script_path = $myinvocation.mycommand.path
$script_folder = Split-Path $script_path -Parent
$bin_debug = Join-Path $script_folder "bin/Debug"
$psd1_path = Join-Path $bin_debug "Visio.psd1"

Resolve-Path $script_folder
Resolve-Path $bin_debug
Resolve-Path $psd1_path

Import-Module $psd1_path
