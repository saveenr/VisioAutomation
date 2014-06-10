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
    -KeepTemporaryFolder $false

Write-Host $result

$choc_filename = Join-Path $output_msi_path ($productshortname + ".nuspec" )
$choc_tools = Join-Path $output_msi_path "tools"

$choc_id = $productshortname
$choc_title = $productname
$choc_ver = "1.3.0"
$choc_authors = $manufacturer
$choc_owners = $manufacturer
$choc_summary = $productname
$choc_description = $productname
$choc_projecturl = $aboutlink
$choc_tags = "Visio PowerShell"
$choc_licenseurl = $aboutlink
$choc_licenseacceptance = "false"
$choc_iconurl = "http://viziblr.com/storage/visioautomation/visioautomation-128x128.png"

$choc_xml = @"
<?xml version="1.0"?>
<package xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <metadata>
    <id>$choc_id</id>
    <title>$choc_title</title>
    <version>$choc_ver</version>
    <authors>$choc_authors</authors>
    <owners>$choc_owners</owners>
    <summary>$choc_summary</summary>
    <description>$choc_description</description>
    <projectUrl>$choc_projecturl</projectUrl>
    <tags>$choc_tags</tags>
    <licenseUrl>$choc_licenseurl</licenseUrl>
    <requireLicenseAcceptance>$choc_licenseacceptance</requireLicenseAcceptance>
    <iconUrl>$choc_iconurl</iconUrl>
  </metadata>
</package>
"@



$choc_xml = [xml] $choc_xml
$choc_xml.Save( $choc_filename )

if (Test-Path $choc_tools)
{
    Remove-Item -Recurse -Force $choc_tools
}
mkdir $choc_tools 

Copy-Item (join-path $output_msi_path "VisioPS_1.2.62.msi") $choc_tools
Copy-Item "D:\saveenr\code\github\visioautomation\VisioAutomation_2010\VisioPowerShell\chocolateyInstall.ps1" $choc_tools
$old = Get-Location
cd $output_msi_path
cpack $choc_filename 
Remove-Item -Recurse -Force $choc_tools
cd $old

Write-Host "----------------------------------------"
Write-Host Done