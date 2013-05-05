function mkdirsafe($path)
{
    if (!(test-path $path))
    {
        mkdir $path
    }
}


$script_path = $myinvocation.mycommand.path
$script_folder = Split-Path $script_path -Parent

Write-Host $script_folder
$bin_debug = Join-Path $script_folder "../bin/debug"
Write-Host $bin_debug

$docfolder =  "$home\documents"
Write-Host $docfolder

$wps =  Join-Path $docfolder "WindowsPowerShell"
Write-Host $wps

$modules =  Join-Path $wps "Modules"
Write-Host $modules 


$visiopsfldr=  Join-Path $modules "VisioPS"
Write-Host $visiopsfldr

mkdirsafe $wps 
mkdirsafe $modules
mkdirsafe $visiopsfldr


Remove-Item $visiopsfldr -Recurse -Force 

mkdir $visiopsfldr | Out-Null

robocopy $bin_debug $visiopsfldr /mir
