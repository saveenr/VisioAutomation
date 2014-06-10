param()

Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"
cls

$scriptpath = Split-Path  $MyInvocation.MyCommand.Path

# ----------------------------------------
# The Most Common User Input settings
#

$productname = "Visio Powershell Module"
$module_foldername = "Visio"
$productshortname = "VisioPS"
$psdfilename = "Visio.psd1"
$manufacturer = "Saveen Reddy"
$helplink = "http://visioautomation.codeplex.com"
$aboutlink = "http://visioautomation.codeplex.com"
$upgradecode = "EE659AB6-BE76-426E-B971-35DF3907F9D4"
$wixbin = Join-path $scriptpath "../../WIX/wix38-binaries"
$mydocs = [Environment]::GetFolderPath("MyDocuments")
$binpath = Resolve-Path ( Join-Path $scriptpath "bin\Debug" )
$output_msi_path = join-path $mydocs ("VisioPSDistribution")
$KeepTempFolderOnExit = $false

if (!(Test-Path $output_msi_path))
{
    mkdir $output_msi_path
}

Write-Host ----------------------------------------

Write-Host "Loading Module Packager PSM1"
$module_packager = Resolve-Path ( Join-Path $scriptpath "PSModulePackager.psm1" )
Import-Module $module_packager


Write-Host "----------------------------------------"
Write-Host Revising PSD1 Version

$Version = Update-PSD1Version

Write-Host "----------------------------------------"
Write-Host Publishing module

$result = Export-PowerShellModuleInstaller `
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
    -IconURL "http://viziblr.com/storage/visioautomation/visioautomation-128x128.png"


Write-Host "----------------------------------------"
Write-Host Done