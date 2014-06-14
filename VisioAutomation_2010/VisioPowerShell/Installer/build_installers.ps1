param()

Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"
cls

$scriptpath = Split-Path  $MyInvocation.MyCommand.Path

# ----------------------------------------
# The Most Common User Input settings

$productname = "Visio Powershell Module"
$module_foldername = "Visio"
$productshortname = "VisioPS"
$psdfilename = "Visio.psd1"
$manufacturer = "Saveen Reddy"
$helplink = "http://visioautomation.codeplex.com"
$aboutlink = "http://visioautomation.codeplex.com"
$upgradecode = "EE659AB6-BE76-426E-B971-35DF3907F9D4"
$wixbin = Join-path $scriptpath "../../../WIX/wix38-binaries"
$mydocs = [Environment]::GetFolderPath("MyDocuments")
$binpath = Resolve-Path ( Join-Path $scriptpath "../bin/Debug" )
$output_msi_path = join-path $mydocs ("VisioPSDistribution")
$KeepTempFolderOnExit = $false
$IconURL = "http://viziblr.com/storage/visioautomation/visioautomation-128x128.png"
$Tags = "Visio PowerShell"

# ----------------------------------------
# Verify Paths

Resolve-Path $scriptpath
Resolve-Path $binpath

if (!(Test-Path $output_msi_path))
{
    mkdir $output_msi_path
}

# ----------------------------------------
# Load Helper Module


$module_packager = Resolve-Path ( Join-Path $scriptpath "PSModulePackager.psm1" )
Import-Module $module_packager

# ----------------------------------------
# Make sure the PSD1 has a new version number


$Old_PSD1 = Join-Path $scriptpath ("../" + $psdfilename )
$New_PSD1 = Join-Path $binpath $psdfilename
$Version = Update-PSD1Version -Old $Old_PSD1 -New $New_PSD1

# ----------------------------------------
# Build the installers

$msi_result = Export-PowerShellModuleInstaller `
    -InputFolder $binpath `
    -ModuleFolderName $module_foldername `
    -OutputFolder $output_msi_path `
    -WIXBinFolder $wixbin `
    -ProductNameLong $productname `
    -ProductNameShort $productshortname `
    -ProductVersion $Version `
    -Manufacturer $manufacturer `
    -ProgramFilesSubFolder $module_foldername `
    -HelpLink $helplink `
    -AboutLink $aboutlink `
    -UpgradeCode $upgradecode `
    -InstallLocationType "PowerShellUserModule" `
    -KeepTemporaryFolder $false `
    -Tags "Visio PowerShell" `
    -IconURL $IconURL `
    -ChocolateyScriptsFolder (Join-Path $scriptpath "Chocolatey") `
    -Verbose 

$msi_result = [PSCustomObject] $msi_result[ $msi_result.Length-1 ] 

$choc_result = Export-ChocolateyPackage `
    -Title $productname `
    -ID $productshortname `
    -Summary "PowerShell module for automation Microsoft Visio 2010 and Visio 2013" `
    -Description "No Description" `
    -Authors "Saveen Reddy" `
    -Owners "Saveen Reddy" `
    -ProjectURL $aboutlink `
    -LicenseURL $aboutlink `
    -ProductVersion $Version `
    -AboutLink $aboutlink `
    -Tags "Visio PowerShell" `
    -IconURL $IconURL `
    -Verbose `
    -MSI $msi_result.MSIFile `

$choc_result = [PSCustomObject] $choc_result[ $choc_result.Length-1 ] 

$msi_result
$choc_result