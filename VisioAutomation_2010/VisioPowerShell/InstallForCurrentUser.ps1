﻿# PURPOSE
# -------
# Manually installs the Visio PowerShell module into the user's PowerShell folder
#
# This is useful when you want to try out the module without going through all the work of 
# creating a module an then installing it with Import-Modile
#
# NOTES
# -----
# - If another PowerShell session has the Visio PS module loaded, then the VisioPS binaries cannot 
#   be replaced by this script because those binaries are locked. In this case, those PS sessions
#   must be terminated before the script will work

Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"

function New-Folder($path)
{
    if (!(test-path $path))
    {
        mkdir $path
    }
}

function Assert-Path($path)
{
    if (!(test-path $path))
    {
		$msg = "Path does not exist " + $path
		Write-Error $msg
    }
}

function Clean-Folder($path)
{
	if (Test-Path $path)
	{
		Remove-Item $path -Recurse -Force 
	}
}

function Mirror-Folder($frompath, $topath)
{
	# /njh - no job header
	# /njh - no job summary
	# /fp - show full paths for files
	# /ns - don't show sizes
	robocopy $frompath $topath /mir /njh /njs /ns /nc /np 
}

function Test-Locked($filePath)
{
    $fileInfo = New-Object System.IO.FileInfo $filePath

    try 
    {
        $fileStream = $fileInfo.Open( [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read )
        return $false
    }
    catch
    {
        return $true
    }
}

# -------------------------------------------
# User-supplied information about this module
$module_foldername = "Visio"
$release = "Debug"

# -------------------
# Calculate the paths

$script_path = $myinvocation.mycommand.path
$script_folder = Split-Path $script_path -Parent
$bin_folder = Join-Path $script_folder ( Join-Path "bin" $release )
$docfolder =  "$home/documents"
$wps =  Join-Path $docfolder "WindowsPowerShell"
$modules =  Join-Path $wps "Modules"
$the_module_folder=  Join-Path $modules $module_foldername
Assert-Path $bin_folder
Assert-Path $docfolder


# ------------------------------
# Verify that the binaries exist
$dlls = Get-ChildItem (Join-Path $bin_folder "*.dll")

if ( ( $dlls -eq $null) -or ( $dlls.Length -lt 1 ) )
{
	$msg = "There are no DLLs in " + $bin_folder
	Write-Error $msg
}	

# Verify key files are There
Assert-Path (Join-Path $bin_folder "Visio.psd1")
Assert-Path (Join-Path $bin_folder "Visio.Types.ps1xml")
Assert-Path (Join-Path $bin_folder "VisioPS.dll")
Assert-Path (Join-Path $bin_folder "VisioScripting.dll")

# ------------------------------
# Prepare the Destination Folder
New-Folder $wps 
New-Folder $modules
Assert-Path $wps
Assert-Path $modules 
Clean-Folder $the_module_folder
New-Folder $the_module_folder
Assert-Path $the_module_folder 

# -----------------
# Copy the contents

Write-Host "Starting mirror"
Write-Host "FROM:" $bin_folder 
Write-Host "TO:" $the_module_folder 

Mirror-Folder $bin_folder $the_module_folder 

Write-Host "Finished!"
