param()

Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"
cls

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

# ----------------------------------------
Write-Host "Loading Module Packager"
$scriptpath = Split-Path  $MyInvocation.MyCommand.Path
$module_packager = Resolve-Path ( Join-Path $scriptpath "PSModulePackager.psm1" )

Import-Module $module_packager

# ----------------------------------------

$mydocs = Get-MyDocsPath
$binpath = Resolve-Path ( Join-Path $scriptpath "bin\Debug" )
$output_msi_path = join-path $mydocs ($productname + " Distribution")
$KeepTempFolderOnExit = $false
$Version = "UNKNOWN"

# ----------------------------------------



Write-Host "----------------------------------------"
Write-Host CREATING updated version number

$Version = Update-PSD1Version

Write-Host "----------------------------------------"
Write-Host Calculating paths, etc.

$baseversion = $Version
$productversion = $baseversion + ".0"
$msibasename = $productshortname + "_" + $baseversion
$output_msi_file = join-path $output_msi_path ($msibasename + ".msi")
$temp_folder = join-path ([Environment]::GetFolderPath("MyDocuments")) ($productshortname  +"_" + (Get-Date -format yyyy_MM_dd))
$cabfilename = $productshortname + ".cab"
$scriptfilename = $MyInvocation.MyCommand.Path
$modules_wxs = join-path $temp_folder ( $productshortname + "_modules.wxs" )
$product_wxs = join-path $temp_folder ( $productshortname + ".wxs" )
$varname = "var." + $productshortname
$wixbin = JoinResolve-path $scriptpath "../../Build/wix36-binaries"
$heatexe = JoinResolve-path $wixbin "heat.exe"
$candleexe = JoinResolve-path $wixbin "candle.exe"
$lightexe = JoinResolve-path $wixbin "light.exe"
$modules_wixobj = join-path $scriptpath ( $productshortname  + "_modules.wixobj" )
$product_wixobj = join-path $scriptpath ( $productshortname + ".wixobj")
$directoryid = $productshortname
$productpdb = join-path (Split-path $output_msi_file) ($msibasename  +".wixpdb")
$licensertf = join-path $binpath "license.rtf"
$licensecmd = ""
$build_file = join-path $binpath "buildinfo.txt"
$build_file_content = "Installer Built on: " + (get-date)
$zipfile = join-path $output_msi_path ($msibasename + ".zip")




if (test-path $licensertf)
{
    $licensecmd = @"
<WixVariable Id="WixUILicenseRtf" Value="License.rtf"></WixVariable>
"@

}


Write-Host "----------------------------------------"
Write-Host CREATING Folders

Remove-FolderIfExists $temp_folder "Temp Folder"
New-Folder $temp_folder
New-Folder $output_msi_path

Write-Host "----------------------------------------"
Write-Host CREATING buildinfo.txt

Set-Content -Value $build_file_content  -Path $build_file


Write-Host "----------------------------------------"
Write-Host "Creating MSI file"

# ----------------------------------------
# - The ProductID is set to zero* because it should be regenerated each time

$modules_xml = @"
<?xml version="1.0" encoding="utf-8"?>
<Wix xmlns='http://schemas.microsoft.com/wix/2006/wi'> 
    <?define RtmProductVersion="1.0.0.0" ?> 
    <?define ProductVersion="#productversion" ?> 

    <Product Id="*" 
		Language="1033" 
		Name="#productname" 
		Version="#productversion"
		Manufacturer="#manufacturer"
		UpgradeCode="#upgradecode">
        <Package Description="#productname Installer" 
		InstallPrivileges="elevated" Comments="#productshortname Installer" 
		InstallerVersion="200" Compressed="yes">
	</Package>
        <Upgrade Id="#upgradecode">
            <UpgradeVersion Minimum="`$(var.ProductVersion)"
                            IncludeMinimum="no"
                            OnlyDetect="yes"
                            Property="NEWPRODUCTFOUND" />
            <UpgradeVersion Minimum="`$(var.RtmProductVersion)"
                            IncludeMinimum="yes"
                            Maximum="`$(var.ProductVersion)"
                            IncludeMaximum="no"
                            Property="UPGRADEFOUND" />
        </Upgrade>

        <!-- Prevent Downgrade -->
        <CustomAction Id="PreventDowngrading" Error="Newer version already installed." />
        <InstallUISequence>
            <Custom Action="PreventDowngrading" After="FindRelatedProducts">NEWPRODUCTFOUND</Custom>
        </InstallUISequence>
        <InstallExecuteSequence>
            <Custom Action="PreventDowngrading" After="FindRelatedProducts">NEWPRODUCTFOUND</Custom>
            <RemoveExistingProducts After="InstallFinalize" />
        </InstallExecuteSequence>
        <Media Id="1" Cabinet="#cabfilename" EmbedCab="yes"></Media>
        #licensecmd 
        <Directory Id="TARGETDIR" Name="SourceDir">
            <Directory Id="PersonalFolder" Name="PersonalFolder">
                <Directory Id="WindowsPowerShell" Name="WindowsPowerShell">
                    <Directory Id="INSTALLDIR" Name="Modules">
                        <Directory Id="#productshortname" Name="#module_foldername">
                        </Directory>
                    </Directory>
                </Directory>
            </Directory>
        </Directory>
        <Property Id="ARPHELPLINK" Value="#helplink"></Property>
        <Property Id="ARPURLINFOABOUT" Value="#aboutlink"></Property>
        <Feature Id="#productshortname" Title="#productshortname" Level="1" ConfigurableDirectory="INSTALLDIR">
            <ComponentGroupRef Id="#productshortname">
            </ComponentGroupRef>
        </Feature>
        <UI></UI>
        <UIRef Id="WixUI_InstallDir"></UIRef>
        <Property Id="WIXUI_INSTALLDIR" Value="INSTALLDIR"></Property>
    </Product>
</Wix>
"@

$modules_xml = $modules_xml -replace "#module_foldername", $module_foldername
$modules_xml = $modules_xml -replace "#productname", $productname
$modules_xml = $modules_xml -replace "#productversion", $productversion
$modules_xml = $modules_xml -replace "#manufacturer", $manufacturer
$modules_xml = $modules_xml -replace "#upgradecode", $upgradecode
$modules_xml = $modules_xml -replace "#productshortname", $productshortname
$modules_xml = $modules_xml -replace "#cabfilename", $cabfilename
$modules_xml = $modules_xml -replace "#licensecmd", $licensecmd
$modules_xml = $modules_xml -replace "#helplink", $helplink
$modules_xml = $modules_xml -replace "#aboutlink", $aboutlink

$modules_xml = [xml] $modules_xml
$modules_xml.Save( $modules_wxs )

# ----------------------------------------
# PRODUCE THE MSI
&$heatexe dir $binpath -nologo -sfrag -suid -ag -srd -dir $directoryid  -out $product_wxs -cg $productshortname  -dr $productshortname -sw5151 -sw5150
&$candleexe $modules_wxs $product_wxs -nologo
&$lightexe -ext WixUIExtension -out $output_msi_file $modules_wixobj $product_wixobj -b $binpath  -sice:ICE91 -sice:ICE69 -sice:ICE38 -sice:ICE57 -sice:ICE64 -sice:ICE204 -nologo

# ----------------------------------------
# CLEANUP 
if (!$KeepTempFolderOnExit)
{
    Remove-FolderIfExists $temp_folder -Description "temp folder"
}

# These have to be manually removed because they don't go into the temp folder by default
Remove-FileIfExists $modules_wixobj "Temp installer file"
Remove-FileIfExists $product_wixobj "Temp installer file"
Remove-FileIfExists $productpdb "Temp installer file"


Write-Host "----------------------------------------"
Write-Host "Creating ZIP file"

Export-ZIPFolder -InputFolder $binpath -OutputFile $zipfile -IncludeBaseDir $false

Write-Host "----------------------------------------"
Write-Host "Script Finished"
Write-Host
Write-Host Files at: $output_msi_path
