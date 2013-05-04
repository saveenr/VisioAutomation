$script_path = $myinvocation.mycommand.path
$script_folder = Split-Path $script_path -Parent

Write-Host $script_folder
$bin_debug = Join-Path $script_folder "bin/debug"
Write-Host $bin_debug

$visiops_dll = Join-Path $bin_debug "VisioPS.dll"
Write-Host $visiops_dll 


$types_file = Join-Path $bin_debug "VisioPS.Types.ps1xml"

Import-Module $visiops_dll 

Update-TypeData $types_file