param([string]$Version="UNKNOWN")

Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"

# ----------------------------------------
# The Most Common User Input settings
#
$productname = "Visio Powershell Module"
$module_foldername = "Visio"
$productshortname = "VisioPS"
$manufacturer = "Saveen Reddy"
$helplink = "http://visioautomation.codeplex.com"
$aboutlink = "http://visioautomation.codeplex.com"
$binpath = (join-path ( Split-Path  $MyInvocation.MyCommand.Path) "bin\Debug")
$binpath = resolve-path $binpath
$upgradecode = "EE659AB6-BE76-426E-B971-35DF3907F9D4"
$output_msi_path = join-path ([Environment]::GetFolderPath("MyDocuments")) ($productname + " Distribution")
$KeepTempFolderOnExit = $false

#------------------------

# MOD version number
$psdfilename = "Visio.psd1"
$src_psd1_filename = Resolve-Path ( Join-Path (Split-Path  $MyInvocation.MyCommand.Path) $psdfilename )
$dst_psd1_filename = Resolve-Path ( Join-Path $binpath $psdfilename )


$src_psd1_filename
$dst_psd1_filename

$psd1_src = Get-Content $src_psd1_filename | Out-String
$psd1_dst = Get-Content $dst_psd1_filename | Out-String


if ( $psd1_src -ne $psd1_dst )
{
    Write-Error "PSD1 files are not the same. Rebuild the project"
    break
}


$psd1_src = Get-Content $src_psd1_filename 
for ($i=0; $i -lt $psd1_src.Length ; $i++)
{
    $src_line = $psd1_src[$i]
    if ($src_line.Trim().StartsWith("ModuleVersion"))
    {
        $tokens = $src_line -split "="
        if ($tokens.Length -ne 2)
        {
            Write-Error "Unexpected number of tokens"
        }

        $tokens2 = $tokens[1].Replace("'","").Trim().Split(".")
        if ($tokens2.Length -ne 3)
        {
            Write-Error "Unexpected number of tokens"
        }

        $lastnum = [int]$tokens2[2]
        $new_lastnum = $lastnum + 1

        $first_num = $tokens2[0]
        $second_num = $tokens2[1]

        $Version = "$first_num.$second_num.$new_lastnum"
        $new_line = "ModuleVersion = '$Version'" 
        Write-Host $new_line
        $psd1_src[$i] = $new_line
    }
}

if ($Version -eq "UNKNOWN")
{
    Write-Error "Version was never set"
}

Write-Host $Version

Set-Content $src_psd1_filename $psd1_src
Set-Content $dst_psd1_filename $psd1_src


# ------------------------------------------

# ----------------------------------------
# Calculate various paths, names, etc baed on user input
# 
$baseversion = $Version
$productversion = $baseversion + "." + (Get-Date -format yyyyMMdd)
$msibasename = $productshortname + "_" + $baseversion
$output_msi_file = join-path $output_msi_path ($msibasename + ".msi")
$temp_folder = join-path ([Environment]::GetFolderPath("MyDocuments")) ($productshortname  +"_" + (Get-Date -format yyyy_MM_dd))
$cabfilename = $productshortname + ".cab"
$scriptfilename = $MyInvocation.MyCommand.Path
$scriptpath = Split-Path $scriptfilename
$modules_wxs = join-path $temp_folder ( $productshortname + "_modules.wxs" )
$product_wxs = join-path $temp_folder ( $productshortname + ".wxs" )
$varname = "var." + $productshortname
$wixbin = join-path $scriptpath "../../Build/wix36-binaries"
Write-Host $wixbin
$wixbin = resolve-path $wixbin
$heatexe = join-path $wixbin "heat.exe"
$candleexe = join-path $wixbin "candle.exe"
$lightexe = join-path $wixbin "light.exe"
$modules_wixobj = join-path $scriptpath ( $productshortname  + "_modules.wixobj" )
$product_wixobj = join-path $scriptpath ( $productshortname + ".wixobj")
$directoryid = $productshortname
$productpdb = join-path (Split-path $output_msi_file) ($msibasename  +".wixpdb")
$licensertf = join-path $binpath "license.rtf"
$licensecmd = ""

if (test-path $output_msi_path )
{
	Write-Verbose "Output Path Exists"
}
else
{
	Write-Verbose "Output Path Does not exists. Creating"
	Write-Host $output_msi_path
	New-Item -Path $output_msi_path -ItemType Directory | Out-Null
}


if (test-path $licensertf)
{
    $licensecmd = @"
<WixVariable Id="WixUILicenseRtf" Value="License.rtf"></WixVariable>
"@

}


# ----------------------------------------
# CREATE THE TEMP FOLDER
if (test-path $temp_folder)
{
    # if it already exists, remote it for safety
    Remove-Item $temp_folder -Recurse
}
New-Item $temp_folder -ItemType directory | Out-Null

$build_file = join-path $binpath "buildinfo.txt"

$build_file_content = "Installer Built on: " + (get-date)

Set-Content -Value $build_file_content  -Path $build_file

# ----------------------------------------
# DYNAMICALLY BUILD THE WXS FILE FOR THE MODULES

# ----------------------------------------
# DYNAMICALLY BUILD THE WXS FILE FOR THE MODULES
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
    if (test-path $temp_folder)
    {
        Remove-Item $temp_folder -Recurse
    }
}

# These have to be manually removed because they don't go into the temp folder by default
remove-item $modules_wixobj 
remove-item $product_wixobj
remove-item $productpdb


#
# Now create the ZIP file


$asm = [Reflection.Assembly]::LoadWithPartialName( "System.IO.Compression.FileSystem" )
$zipfile = join-path $output_msi_path ($msibasename + ".zip")

if (test-path $zipfile)
{
	remove-item $zipfile
}

$compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
$includebasedir = $false
[System.IO.Compression.ZipFile]::CreateFromDirectory($binpath ,$zipfile ,$compressionLevel, $includebasedir )

Write-Host "Script Finished"
Write-Host MSI and ZIP file stored at: $output_msi_path
