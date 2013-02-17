Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"

$script_path = $myinvocation.mycommand.path
$script_folder = Split-Path $script_path -Parent
$project_path = Split-Path $script_folder -Parent
$bindebug_path = Join-Path $src_path "bin/Debug"
$localdll = Join-Path $bindebug_path "VisioPS.dll"

if (!(Test-Path $visiops_localdll))
{
    Write-Host "Error: Cannot find" $localdll
    Break
}

Import-Module $visiops_localdll 

$cmds = Get-Command -Module VisioPS

foreach ($cmd in $cmds) 
{ 
    Write-Host "--------------------"
    Write-Host

	get-help $cmd.Name 
} 

