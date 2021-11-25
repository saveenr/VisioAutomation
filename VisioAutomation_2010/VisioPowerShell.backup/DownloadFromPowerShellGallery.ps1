﻿# Loads the module from bin/debug folder.

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
