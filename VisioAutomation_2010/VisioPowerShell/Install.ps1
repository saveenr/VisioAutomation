# Installs this module for the User
# This is for when you want to quickly check that normal installed usage works
# but don't want to go through the full process of generating the installer, etc.
# NOTE: If another PS Session has the module loaded the binaries cannot be replaced

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

# ------------------------------
# Prepare the Distination Folder
New-Folder $wps 
New-Folder $modules
Assert-Path $wps
Assert-Path $modules 
Clean-Folder $the_module_folder
New-Folder $the_module_folder
Assert-Path $the_module_folder 

# -----------------
# Copy the contents
Mirror-Folder $bin_folder $the_module_folder 

