# Make_PSMod_MSI
# This script makes it easy to create an MSI to install a simple powershell module
# By: Saveen Reddy
#
# Inspired by : http://sev17.com/2010/11/building-a-powershell-module-installer/
# More information on WIX: http://themech.net/2008/08/how-to-create-a-simple-msi-installer-using-wix/

Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"


# ----------------------------------------
# USER INPUTS TO THE BUILD PROCESS
$productname = "Visio Powershell Module"
$productshortname = "VisioPS"
$productversion = "1.0.0." + (Get-Date -format yyyyMMdd)
$manufacturer = "Saveen Reddy"
$helplink = "http://visioautomation.codeplex.com"
$aboutlink = "http://visioAutomation.codeplex.com"
$binpath = "D:\saveenr\code\visioautomation\VisioAutomation_2010\VisioPS\bin\Debug"
$productid = "4A2B528A-93E5-431D-97BB-79767C7677C5"
$upgradecode = "EE659AB6-BE76-426E-B971-35DF3907F9D4"
$upgradeid = "F14DC5AF-1234-498A-9646-AA27E03957AA"
$output_msi_folder = [Environment]::GetFolderPath("MyDocuments")

Write-Host Source Folder: $binpath
Write-Host MSI Will be placed here: $output_msi_folder
Write-Host

# ----------------------------------------
# CALCULATE VARIOUS PATHS, FILENAMES, IDS, BASED ON INPUT
$delete_temp_folder_on_exit = $true;
$datestring = Get-Date -format yyyy-MM-dd
$temp_folder = join-path ([Environment]::GetFolderPath("MyDocuments")) ($productshortname  +"_" + $datestring)
$cabfilename = $productshortname + ".cab"
$scriptfilename = $MyInvocation.MyCommand.Path
$scriptpath = Split-Path $scriptfilename
$modules_wxs = join-path $temp_folder ( $productshortname + "_modules.wxs" )
$product_wxs = join-path $temp_folder ( $productshortname + ".wxs" )
$varname = "var." + $productshortname
$wixbin = join-path $scriptpath "wix36-binaries"
$heatexe = join-path $wixbin "heat.exe"
$candleexe = join-path $wixbin "candle.exe"
$lightexe = join-path $wixbin "light.exe"
$modules_wixobj = join-path $scriptpath ( $productshortname  + "_modules.wixobj" )
$product_wixobj = join-path $scriptpath ( $productshortname + ".wixobj")
$directoryid = $productshortname
$msibasename = $productshortname + "_" + (Get-Date -format yyyyMMdd)
$output_msi_file = join-path $output_msi_folder ($msibasename + ".msi")
$productpdb = join-path (Split-path $output_msi_file) ($msibasename +".wixpdb")
$licensertf = join-path $binpath "license.rtf"
$licensecmd = ""


# ----------------------------------------
# if a License.rtf file exists use it
if (test-path $licensertf)
{
    $licensecmd = @"
<WixVariable Id="WixUILicenseRtf" Value="License.rtf"></WixVariable>
"@
}

# ----------------------------------------
# CREATE BEFORE WE BEGIN
if (test-path $temp_folder)
{
    # if it already exists, remote it for safety
    Remove-Item $temp_folder -Recurse
}
New-Item $temp_folder -ItemType directory | Out-Null

if (test-path $productpdb)
{
    Remove-Item $productpdb
}

if (test-path $output_msi_file)
{
    Remove-Item $output_msi_file
}

# --------------d:\--------------------------
# VALIDATE THE BINARIES EXIST
if (!(test-path $heatexe))
{
    Write-Error Could not find $heatexe
}
if (!(test-path $candleexe))
{
    Write-Error Could not find $candleexe
}
if (!(test-path $lightexe))
{
    Write-Error Could not find $lightexe
}


# ----------------------------------------
# DYNAMICALLY BUILD THE WXS FILE FOR THE MODULES
$modules_xml = [xml] @"
<?xml version="1.0" encoding="utf-8"?>
<Wix xmlns='http://schemas.microsoft.com/wix/2006/wi'> 
    <Product Id="$productid" 
		Language="1033" 
		Name="$productname" 
		Version="$productversion"
		Manufacturer="$manufacturer"
		UpgradeCode="$upgradecode">
        <Package Description="$productname Installer" 
		InstallPrivileges="elevated" Comments="$productshortname Installer" 
		InstallerVersion="200" Compressed="yes">
	</Package>
        <Upgrade Id="$upgradeid">
            <UpgradeVersion 
		        OnlyDetect="no" 
		        Property="PREVIOUSFOUND" 
		        Minimum="1.0.0" 
		        IncludeMinimum="yes" 
		        Maximum="1.0.0"
		        IncludeMaximum="no">
        	</UpgradeVersion>
        </Upgrade>
        <InstallExecuteSequence>
            <RemoveExistingProducts After="InstallInitialize"></RemoveExistingProducts>
        </InstallExecuteSequence>
        <Media Id="1" Cabinet="$cabfilename" EmbedCab="yes"></Media>
        $licensecmd 
        <Directory Id="TARGETDIR" Name="SourceDir">
            <Directory Id="PersonalFolder" Name="PersonalFolder">
                <Directory Id="WindowsPowerShell" Name="WindowsPowerShell">
                    <Directory Id="INSTALLDIR" Name="Modules">
                        <Directory Id="$productshortname" Name="$productshortname">
                        </Directory>
                    </Directory>
                </Directory>
            </Directory>
        </Directory>
        <Property Id="ARPHELPLINK" Value="$helplink"></Property>
        <Property Id="ARPURLINFOABOUT" Value="$aboutlink"></Property>
        <Feature Id="$productshortname" Title="$productshortname" Level="1" ConfigurableDirectory="INSTALLDIR">
            <ComponentGroupRef Id="$productshortname">
            </ComponentGroupRef>
        </Feature>
        <UI></UI>
        <UIRef Id="WixUI_InstallDir"></UIRef>
        <Property Id="WIXUI_INSTALLDIR" Value="INSTALLDIR"></Property>
    </Product>
</Wix>
"@


# ----------------------------------------
# PRODUCE THE WXS FILES
$modules_xml.Save( $modules_wxs )
if (!(test-path $modules_wxs ))
{
    Write-Error Could not find  $modules_wxs
}

&$heatexe dir $binpath -nologo -sfrag -suid -ag -srd -dir $directoryid  -out $product_wxs -cg $productshortname  -dr $productshortname
if (!(test-path $product_wxs ))
{
    Write-Error Did not produce $product_wxs
}


# ----------------------------------------
# PRODUCE THE WIXOBJ FILES
&$candleexe $modules_wxs $product_wxs 
if (!(test-path $modules_wixobj ))
{
    Write-Error Did not produce $modules_wixobj 
}
if (!(test-path $product_wixobj ))
{
    Write-Error Did not produce $product_wixobj
}

# ----------------------------------------
# PRODUCE THE MSI

&$lightexe -ext WixUIExtension -out $output_msi_file $modules_wixobj $product_wixobj -b $binpath -sice:ICE91 -sice:ICE69 -sice:ICE38 -sice:ICE57 -sice:ICE64 -sice:ICE204
if (!(test-path $productpdb))
{
    Write-Error Did not produce $productpdb
}

if (!(test-path $output_msi_file))
{
    Write-Error Did not produce $output_msi_file
}


# ----------------------------------------
# CLEANUP 
if ($delete_temp_folder_on_exit)
{
    if (test-path $temp_folder)
    {
        Remove-Item $temp_folder -Recurse
    }
}

# These have to be manually removed because they don't go into the temp folder by default
remove-item $modules_wixobj 
remove-item $product_wixobj
remove-item $productpdb


# ----------------------------------------
# FINAL MESSAGE
Write-Host 
Write-Host "----------------------------------------"
Write-Host "SUCCESS"
Write-Host MSI Created: $output_msi_file 

