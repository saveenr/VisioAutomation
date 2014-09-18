# CodePackage
# Cmdlets to help distribute code
#
# HISTORY
#
# 2014-06-13 - Version 1.3
# - Added Chocolatey integration
#
# 2013-03-11 - Version 1.2
# - Added cmdlet to cleanly install a powershell module
#
# 2013-03-08 - Version 1.1
# - Now MSI filename uses version instead of datestring
# - Fixed Bug in Temp Folder Deletion
#
# 2013-02-23 - Version 1.0
# - Initial version

Set-StrictMode -Version 2 
$ErrorActionPreference = "Stop"

function Get-MyDocsPath()
{
    [Environment]::GetFolderPath("MyDocuments")
}

function Test-TextFilesAreEqual( $left, $right )
{
    $left_text = Get-Content $left | Out-String
    $right_text = Get-Content $right | Out-String
    ($left_text -eq $right_text)
}

function New-Folder
{
	param
	(
		[parameter(Mandatory=$true)] [string] $Folder
	)
	process
	{
        if (test-path $Folder )
        {
        }
        else
        {
	        New-Item -Path $Folder -ItemType Directory | Out-Null
        }

	}
}

function Remove-Folder
{
	param
	(
		[parameter(Mandatory=$true)] [string] $Folder
	)
	process
	{
		if ( test-path $Folder)
		{
    		Remove-Item $Folder -Recurse
		}
	}
}

function Remove-File
{
	param
	(
		[parameter(Mandatory=$true)] [string] $Filename
	)
	process
	{
		if ( test-path $Filename)
		{
    		Remove-Item $Filename -Recurse
		}
	}
}


function Copy-CodeFolder
{
	param
	(
		[parameter(Mandatory=$true)] [string] $SourceFolder , 
		[parameter(Mandatory=$true)] [string] $OutputFolder
	)
	process
	{
	
		# ---------------------------------
		# COPY FILES TO THE STAGING FOLDER
		# Remove the read-only flag with /A-:R
		# Exclude Files with /XF option
		#  *.suo 
		#  *.user 
		#  *.vssscc 
		#  *.vspscc 
		# Exclude folders with /XD option
		#  bin
		#  obj
		#  _Resharper

		# Control verbosity 
		#  Don't show the names of files /NFL
		#  Don't show the names of directories /NDL
		&robocopy $SourceFolder $OutputFolder /MIR /A-:R /XF *.suo /XF *.user /XF *.vssscc /XF *.vspscc /XF *.ignore /XF *.temp /XF *.tmp /NFL /NDL /XD bin /XD obj /XD _ReSharper*
	}
}


function AssertPathExists( $p )
{
    Write-Verbose "Checking path exists $p"
    if (Test-Path $p)
    {
    }
    else
    {
        $msg = "Path does not exist: $p"
        $exc = New-Object System.ArgumentException $msg
        Throw $exc
    }
}

function AssertFileExists( $p )
{
    Write-Verbose "Checking file exists $p"
    if (Test-Path $p)
    {
    }
    else
    {
        $msg = "File does not exist"
        $exc = New-Object System.ArgumentException $msg
        Throw $exc
    }
}

function AssertFileWasProduced( $p )
{
    Write-Verbose "Checking file was produced $p"
    if (Test-Path $p)
    {
    }
    else
    {
        $msg = "File was not produced"
        $exc = New-Object System.ArgumentException $msg
        Throw $exc
    }
}


function Export-PowerShellModuleInstaller
{
    param (
        [parameter(Mandatory=$true)] [string] $InputFolder,
        [parameter(Mandatory=$true)] [string] $OutputFolder,
        [parameter(Mandatory=$true)] [string] $WIXBinFolder,
        [parameter(Mandatory=$true)] [string] $ProductNameLong,
        [parameter(Mandatory=$true)] [string] $ProductNameShort,
        [parameter(Mandatory=$true)] [string] $ProductVersion,
        [parameter(Mandatory=$true)] [string] $Manufacturer,
        [parameter(Mandatory=$true)] [string] $ModuleFolderName,
        [parameter(Mandatory=$true)]
		[AllowEmptyString()]
		[string] $ProgramFilesSubFolder,
        [parameter(Mandatory=$true)] [string] $HelpLink,
        [parameter(Mandatory=$true)] [string] $AboutLink,
        [parameter(Mandatory=$true)] [string] $UpgradeCode,
        [parameter(Mandatory=$true)] 
		[ValidateSet("Default","ProgramFiles","PowerShellUserModule")] 
		[string] $InstallLocationType,
        [parameter(Mandatory=$false)] [bool] $KeepTemporaryFolder,
        [parameter(Mandatory=$true)] [string] $Tags,
        [parameter(Mandatory=$true)] [string] $IconURL,
        [parameter(Mandatory=$true)] [string] $ChocolateyScriptsFolder
		
    )
    PROCESS 
    {
        $ProductID = "*" # so it regenerate devery time
        $UpgradeID = $UpgradeCode # so it regenerate devery time
        # ----------------------------------------
        # VERIFY USER INPUT
        AssertPathExists( $InputFolder )
        AssertPathExists( $WIXBinFolder )
        AssertPathExists( $OutputFolder )

        # ----------------------------------------
        # CALCULATE VARIOUS PATHS, FILENAMES, IDS, BASED ON INPUT

        $datestring = Get-Date -format yyyy-MM-dd
        $temp_folder = join-path ([Environment]::GetFolderPath("MyDocuments")) ($ProductNameShort  +"_" + $datestring)
        $cabfilename = $ProductNameShort + ".cab"
        $modules_wxs = join-path $temp_folder ( $ProductNameShort + "_modules.wxs" )
        $product_wxs = join-path $temp_folder ( $ProductNameShort + ".wxs" )
        $varname = "var." + $ProductNameShort
        $heatexe = join-path $WIXBinFolder "heat.exe"
        $candleexe = join-path $WIXBinFolder "candle.exe"
        $lightexe = join-path $WIXBinFolder "light.exe"
        $modules_wixobj = join-path (Get-Location) ( $ProductNameShort  + "_modules.wixobj" )
        $product_wixobj = join-path (Get-Location) ( $ProductNameShort + ".wixobj")
        $directoryid = $ProductNameShort
        $msibasename = $ProductNameShort + "_" + $ProductVersion
        $msifilename =  $msibasename + ".msi"
        $output_msi_file = join-path $OutputFolder $msifilename
        $productpdb = join-path (Split-path $output_msi_file) ($msibasename +".wixpdb")
        $licensertf = join-path $InputFolder "license.rtf"

        if (test-path $licensertf)
        {
            $licensecmd = "<WixVariable Id=`"WixUILicenseRtf`" Value=`"License.rtf`"></WixVariable>"
        }
        else
        {
            $licensecmd = ""
        }

        # ----------------------------------------
        # CREATE BEFORE WE BEGIN
        if (test-path $temp_folder)
        {
            # if it already exists, remote it for safety
            Remove-Item $temp_folder -Recurse
        }
        New-Item $temp_folder -ItemType directory | Out-Null

        Remove-File $productpdb
        Remove-File $output_msi_file

        # --------------d:\--------------------------
        # VALIDATE THE BINARIES EXIST
        AssertFileExists($heatexe)
        AssertFileExists($candleexe)
        AssertFileExists($lightexe)


	    # ----------------------------------------
	    # DYNAMICALLY BUILD THE WXS FILE FOR THE MODULES

        $installdir = $null

        $powershell_user_module_installdir = @"
<Directory Id="PersonalFolder" Name="PersonalFolder">
    <Directory Id="WindowsPowerShell" Name="WindowsPowerShell">
        <Directory Id="INSTALLDIR" Name="Modules">
            <Directory Id="$ProductNameShort" Name="$ModuleFolderName">
            </Directory>
        </Directory>
    </Directory>
</Directory>
"@

        $program_files_installdir = @"
<Directory Id="ProgramFilesFolder">
        <Directory Id="INSTALLDIR" Name="$ProgramFilesSubFolder">
            <Directory Id="$ProductNameShort" Name="$ProductNameShort">
            </Directory>
        </Directory>
</Directory>
"@

		# this has to be done first
		if ($InstallLocationType -eq "PowerShellUserModule")
		{
            $installdir = $powershell_user_module_installdir
		}
		elseif( ($InstallLocationType -eq "Default") -or ($InstallLocationType -eq "ProgramFiles"))
		{
			if ( ($ProgramFilesSubFolder -eq $null) -or ($ProgramFilesSubFolder -eq ""))
			{
                $msg = "$ProgramFilesSubFolder is null or empty"
                $exc = New-Object System.ArgumentException $msg
                Throw $exc
			}
             $installdir = $program_files_installdir
		}
		else
		{
            $msg = "Unsupported InstallType $InstallLocationType "
            $exc = New-Object System.ArgumentException $msg
            Throw $exc
		}


	    $modules_xml = @"
<?xml version="1.0" encoding="utf-8"?>
<Wix xmlns='http://schemas.microsoft.com/wix/2006/wi'> 
    <Product Id="$ProductID" 
		Language="1033" 
		Name="$ProductNameLong" 
		Version="$ProductVersion"
		Manufacturer="$Manufacturer"
		UpgradeCode="$UpgradeCode">
        <Package Description="$ProductNameLong Installer" 
		InstallPrivileges="elevated" Comments="$ProductNameShort Installer" 
		InstallerVersion="200" Compressed="yes">
	</Package>
        <Upgrade Id="$UpgradeID">
            <UpgradeVersion 
		        OnlyDetect="no" 
		        Property="PREVIOUSFOUND" 
		        Minimum="1.0.0" 
		        IncludeMinimum="yes" 
		        Maximum="1.0.0.0"
		        IncludeMaximum="no">
        	</UpgradeVersion>
        </Upgrade>
        <InstallExecuteSequence>
            <RemoveExistingProducts After="InstallInitialize"></RemoveExistingProducts>
        </InstallExecuteSequence>
        <Media Id="1" Cabinet="$cabfilename" EmbedCab="yes"></Media>
        $licensecmd
        <Directory Id="TARGETDIR" Name="SourceDir">
		$installdir
        </Directory>
        <Property Id="ARPHELPLINK" Value="$HelpLink"></Property>
        <Property Id="ARPURLINFOABOUT" Value="$AboutLink"></Property>
        <Feature Id="$ProductNameShort" Title="$ProductNameShort" Level="1" ConfigurableDirectory="INSTALLDIR">
            <ComponentGroupRef Id="$ProductNameShort">
            </ComponentGroupRef>
        </Feature>
        <UI></UI>
        <UIRef Id="WixUI_InstallDir"></UIRef>
        <Property Id="WIXUI_INSTALLDIR" Value="INSTALLDIR"></Property>
    </Product>
</Wix>
"@

	    $modules_xml = [xml] $modules_xml
	    $modules_xml.Save( $modules_wxs )
		
        # ----------------------------------------
        # PRODUCE THE WXS FILE
        $modules_xml.Save( $modules_wxs )
        AssertFileWasProduced( $modules_wxs )

        &$heatexe dir $InputFolder -nologo -sfrag -suid -ag -srd -dir $directoryid  -out $product_wxs -cg $ProductNameShort  -dr $ProductNameShort
        AssertFileWasProduced( $product_wxs )

        # ----------------------------------------
        # PRODUCE THE WIXOBJ FILES VIA CANDLE
        &$candleexe $modules_wxs $product_wxs 
        AssertFileWasProduced( $modules_wixobj )
        AssertFileWasProduced( $product_wixobj )

        # ----------------------------------------
        # PRODUCE THE MSI VIA LIGHT
        &$lightexe -ext WixUIExtension -out $output_msi_file $modules_wixobj $product_wixobj -b $InputFolder -sice:ICE91 -sice:ICE69 -sice:ICE38 -sice:ICE57 -sice:ICE64 -sice:ICE204
        AssertFileWasProduced($productpdb)
        AssertFileWasProduced($output_msi_file)

        # ----------------------------------------
        # CLEANUP 
        if (!($KeepTemporaryFolder))
        {
			Remove-Folder -Folder $temp_folder -Verbose
        }

        # These have to be manually removed because they don't go into the temp folder by default
        Remove-File $modules_wixobj 
        Remove-File $product_wixobj 
        Remove-File $productpdb 

        Write-Verbose "Creating ZIP file"
        $zipfile = join-path $OutputFolder ($msibasename + ".zip")
        Export-ZIPFolder -InputFolder $InputFolder -OutputFile $zipfile -IncludeBaseDir $false
        

        $result = [PSCustomObject] @{ MSIFile = $output_msi_file ; ZipFile = $zipfile; ProductVersion = $ProductVersion }
        $result
    }
}

function Export-ChocolateyPackage
{
    param (
        [parameter(Mandatory=$true)] [string] $Title,
        [parameter(Mandatory=$true)] [string] $ID,
        [parameter(Mandatory=$true)] [string] $Summary,
        [parameter(Mandatory=$true)] [string] $Description,
        [parameter(Mandatory=$true)] [string] $Authors,
        [parameter(Mandatory=$true)] [string] $Owners,
        [parameter(Mandatory=$true)] [string] $ProductVersion,
        [parameter(Mandatory=$true)] [string] $ProjectURL,
        [parameter(Mandatory=$true)] [string] $LicenseURL,
        [parameter(Mandatory=$true)] [string] $AboutLink,
        [parameter(Mandatory=$true)] [string] $Tags,
        [parameter(Mandatory=$true)] [string] $IconURL,
        [parameter(Mandatory=$true)] [string] $MSI
		
    )
    PROCESS 
    {

        $MSI = Resolve-Path $MSI
        $OutputFolder = Resolve-Path (Split-Path $MSI)

        # http://www.topbug.net/blog/2012/07/02/a-simple-tutorial-create-and-publish-chocolatey-packages/

        Write-Verbose "Creating Chocolatey package"
        $choc_filename = Join-Path $OutputFolder ($ID + ".nuspec" )
        $choc_licenseacceptance = "false"
                
        $choc_msi_file = Split-Path $MSI -Leaf

        $choc_xml = @"
<?xml version="1.0"?>
<package xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <metadata>
    <id>$ID</id>
    <title>$Title</title>
    <version>$ProductVersion</version>
    <authors>$Authors</authors>
    <owners>$Owners</owners>
    <summary>$Summary</summary>
    <description>$Description</description>
    <projectUrl>$ProjectURL</projectUrl>
    <tags>$Tags</tags>
    <licenseUrl>$LicenseURL</licenseUrl>
    <requireLicenseAcceptance>$choc_licenseacceptance</requireLicenseAcceptance>
    <iconUrl>$IconURL</iconUrl>
  </metadata>
  <files>
    <file src="$choc_msi_file" target="tools" />
    <file src="chocolateyInstall.ps1" target="tools" />
  </files>
</package>
"@

        $choc_installps1text= @"
`$packageName = "$ID"
`$installerType = 'MSI' 
`$url = "$choc_msi_file" 
`$silentArgs = '/quiet' 
`$validExitCodes = @(0) 

Install-ChocolateyPackage "`$packageName" "`$installerType" "`$silentArgs" "`$url" -validExitCodes `$validExitCodes
"@

        $choc_xml = [xml] $choc_xml
        $choc_xml.Save( $choc_filename )


        Write-Verbose "Populating Chocolately Tools directory"
        $choc_installps1text | Out-File (Join-Path $OutputFolder "chocolateyInstall.ps1")

        $old = Get-Location
        Resolve-Path $OutputFolder

        $pkg_filename = $ID + "." + $ProductVersion + ".nupkg"
        $choc_pkg = Join-Path $OutputFolder $pkg_filename
        Remove-File $choc_pkg

        Resolve-Path $OutputFolder
        try
        {
            Write-Verbose "Changing location to $OutputFolder"
            cd $OutputFolder

            Write-Verbose "Cleaning Chocolately package"
            $choc_results = cpack $choc_filename -Verbose
        }
        Finally
        {
            cd $old
        }

        AssertFileExists $choc_pkg

        $choc_test_script= "cinst $ID -Source %cd%";
        $cmd = Join-Path  $OutputFolder "InstallChocolateyPackageLocal.cmd"
        $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding($False)
        [System.IO.File]::WriteAllLines($cmd , $choc_test_script, $Utf8NoBomEncoding)

        $result = [PSCustomObject] @{ 
            MSIFile = $MSI ; 
            ChocolateyPackage = $choc_pkg; 
            ProductVersion = $ProductVersion }
        $result

    }
}



function Export-ZIPFolder
{
    param (
        [parameter(Mandatory=$true)] [string] $InputFolder,
        [parameter(Mandatory=$true)] [string] $OutputFile,
        [parameter(Mandatory=$true)] [bool] $IncludeBaseDir,
        [parameter(Mandatory=$false)] [switch] $Overwrite
    )
    PROCESS 
    {
        if (Test-Path $OutputFile)
        {
            Remove-Item $OutputFile
        }

        $asm = [Reflection.Assembly]::LoadWithPartialName( "System.IO.Compression.FileSystem" )
        $compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
        [System.IO.Compression.ZipFile]::CreateFromDirectory($InputFolder, $OutputFile, $compressionLevel, $IncludeBaseDir )
    }
}

function Remove-InstalledProgram
{
    param (
        [parameter(Mandatory=$true)] [string] $Name
    )
    PROCESS 
    {
        $filter = "Name = '$Name'"
        $app = Get-WmiObject -Class Win32_Product -Filter $filter
        if ($app -eq $null)
        {
            Write-Verbose "Program not installed"
        }
        else
        {
            Write-Verbose "App is installed"
            Write-Verbose "Uninstalling now"
            $app.Uninstall()
            Write-Verbose "Finished Uninstalling"   
        }
    }
}

function Install-PSModuleFromFolder
{
    param (
        [parameter(Mandatory=$true)] [string] $Folder,
        [parameter(Mandatory=$true)] [string] $Name
    )
    PROCESS 
    {
        $mydocs = join-Path $env:USERPROFILE Documents
        $output_folder = Join-Path $mydocs "WindowsPowerShell\Modules"
        $output_folder = Join-Path $output_folder $Name
        
        if (Test-Path $output_folder)
        {
            foreach ($i in Get-ChildItem $output_folder)
            {
                $fn =  $i.FullName
                Write-Verbose "Removing $fn"
                Remove-Item $i.FullName -Recurse -Force
            }
        }
        else
        {        [parameter(Mandatory=$true)] [string] $ScriptPath

            New-Item $output_folder -ItemType Directory
        }

        &robocopy $Folder $output_folder /MIR /A-:R /XF *.pdb /XF *.ignore 
    }
}

function Update-Version($old_version, $index)
{
    $tokens2 = $old_version.Split(".")

    $lastnum = [int]$tokens2[$index]
    $new_lastnum = $lastnum + 1

    $first_num = $tokens2[0]
    $second_num = $tokens2[1]

    $new_version = "$first_num.$second_num.$new_lastnum"
    $new_version
}


function Update-PSD1Version
{
    param (
        [parameter(Mandatory=$true)] [string] $Old,
        [parameter(Mandatory=$true)] [string] $New
    )
    PROCESS 
    {

        $src_psd1_filename = $Old
        $dst_psd1_filename = $New

        if (!( Test-TextFilesAreEqual $src_psd1_filename $dst_psd1_filename))
        {
            $exc = New-Object System.ArgumentException "PSD1 files are not the same. Rebuild the project"
            Throw $exc
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
                    $msg = "Unexpected number of tokens"
                    $exc = New-Object System.ArgumentException $msg
                    Throw $exc
                }

                $old_version = $tokens[1].Replace("'","").Trim()

                $Version = Update-Version $old_version 2
                $new_line = "ModuleVersion = '$Version'" 
                $psd1_src[$i] = $new_line
            }
        }

        if ($Version -eq "UNKNOWN")
        {
            $msg = Write-Error "Version was never set"
            $exc = New-Object System.ArgumentException $msg
            Throw $exc
        }

        Set-Content $src_psd1_filename $psd1_src
        Set-Content $dst_psd1_filename $psd1_src

        $Version
    }
}

