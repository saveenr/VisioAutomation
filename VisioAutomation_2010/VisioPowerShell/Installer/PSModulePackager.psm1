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
        $output_msi_file = join-path $OutputFolder ($msibasename + ".msi")
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
        $zipfile = join-path $output_msi_path ($msibasename + ".zip")
        Export-ZIPFolder -InputFolder $binpath -OutputFile $zipfile -IncludeBaseDir $false
        

        # ---------------------------------------
        # CHOCOLATEY
        # http://www.topbug.net/blog/2012/07/02/a-simple-tutorial-create-and-publish-chocolatey-packages/
        # cinst .\VisioPS.N.N.N.nupkg -Source Get-Location

        Write-Verbose "Creating Chocolatey package"
        $choc_filename = Join-Path $OutputFolder ($productshortname + ".nuspec" )
        $choc_tools = Join-Path $OutputFolder "tools"

        $choc_id = $ProductNameShort
        $choc_title = $ProductNameLong
        $choc_ver = $ProductVersion
        $choc_authors = $Manufacturer
        $choc_owners = $Manufacturer
        $choc_summary = $ProductNameLong
        $choc_description = $ProductNameLong
        $choc_projecturl = $AboutLink
        $choc_tags = $Tags
        $choc_licenseurl = $AboutLink
        $choc_licenseacceptance = "false"
        $choc_iconurl = $IconURL

        $choc_msi_file = ($msibasename + ".msi")

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
  <files>
    <file src="$choc_msi_file" target="content" />
    <file src="tools\chocolateyInstall.ps1" target="tools" />
  </files>
</package>
"@

        $choc_installps1text= @"
`$packageName = "$ProductNameLong"
`$installerType = 'MSI' 
`$url = "$choc_msi_file" 
`$silentArgs = '/quiet' 
`$validExitCodes = @(0) 

Install-ChocolateyPackage "`$packageName" "`$installerType" "`$silentArgs" "`$url" -validExitCodes `$validExitCodes
"@

        $choc_xml = [xml] $choc_xml
        $choc_xml.Save( $choc_filename )

        if (Test-Path $choc_tools)
        {
            Remove-Item -Recurse -Force $choc_tools
        }

        Write-Verbose "Populating Chocolately Tools directory"
        mkdir $choc_tools 
        $choc_install_script = Join-Path $ChocolateyScriptsFolder "chocolateyInstall.ps1"
        Copy-Item $output_msi_file $choc_tools 
        #Copy-Item $choc_install_script $choc_tools
        $choc_installps1text | Out-File (Join-Path $choc_tools "chocolateyInstall.ps1")


        $old = Get-Location
        Resolve-Path $OutputFolder

        $choc_pkg = Join-Path $OutputFolder ($ProductNameShort + "." + $ProductVersion + ".nupkg")
        Remove-File $choc_pkg

        try
        {
            Write-Verbose "Changing location to $OutputFolder"
            Resolve-Path $OutputFolder
            cd $OutputFolder

            Write-Verbose "Cleaning Chocolately package"
            $choc_results = cpack $choc_filename -Verbose
        }
        Finally
        {
            cd $old
        }

        AssertFileExists $choc_pkg

        Write-Verbose "Cleaning Chocolately Tools directory"
        #Remove-Item -Recurse -Force $choc_tools

        $result = [PSCustomObject] @{ MSIFile = $output_msi_file ; ZipFile = $zipfile; ProductVersion = $ProductVersion }
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

