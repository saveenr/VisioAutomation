# CodePackage
# Cmdlets to help distribute code

# HISTORY
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

function JoinResolve-Path($a, $b)
{
    $p = Join-Path $a $b
    Write-Host $p
    Resolve-Path $p
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

function Remove-FolderIfExists
{
	param
	(
		[parameter(Mandatory=$true)] [string] $Folder , 
		[parameter(Mandatory=$true)] [string] $Description 
	)
	process
	{
		if ( test-path $Folder)
		{
			Write-Verbose $Description
    		Remove-Item $Folder -Recurse
		}
	}
}

function Remove-FileIfExists
{
	param
	(
		[parameter(Mandatory=$true)] [string] $Filename , 
		[parameter(Mandatory=$true)] [string] $Description 
	)
	process
	{
		if ( test-path $Filename)
		{
			Write-Verbose $Description
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
        Write-Host "ERROR: Path does not exist"
        Break    
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
        Write-Error "ERROR: File does not exist"
        Break    
    }
}

function AssertFileWasProduced( $p )
{
    Write-Host "Checking file was produced" $p
    if (Test-Path $p)
    {
    }
    else
    {
        Write-Host "ERROR: File was not produced"
        Break    
    }
}

function DeleteFileIfExists( $filename )
{
    if (test-path $filename )
    {
        Remove-Item $filename 
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
        [parameter(Mandatory=$false)] [bool] $KeepTemporaryFolder
		
    )
    PROCESS 
    {
        $ProductID = "*" # so it regenerate devery time
        $UpgradeID = $UpgradeCode # so it regenerate devery time
        # ----------------------------------------
        # VERIFY USER INPUT
        Write-Host 
        Write-Host Veryify Paths
        AssertPathExists( $InputFolder )
        AssertPathExists( $WIXBinFolder )
        AssertPathExists( $OutputFolder )
        Write-Host Finished Veryifying Paths

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


        Write-Host Source Folder to Package: $InputFolder
        Write-Host MSI Will be placed here: $OutputFolder
        Write-Host

        # ----------------------------------------
        # CREATE BEFORE WE BEGIN
        if (test-path $temp_folder)
        {
            # if it already exists, remote it for safety
            Remove-Item $temp_folder -Recurse
        }
        New-Item $temp_folder -ItemType directory | Out-Null

        DeleteFileIfExists($productpdb)
        DeleteFileIfExists($output_msi_file)

        # --------------d:\--------------------------
        # VALIDATE THE BINARIES EXIST
        AssertFileExists($heatexe)
        AssertFileExists($candleexe)
        AssertFileExists($lightexe)


	    # ----------------------------------------
	    # DYNAMICALLY BUILD THE WXS FILE FOR THE MODULES
	    $modules_xml = @"
<?xml version="1.0" encoding="utf-8"?>
<Wix xmlns='http://schemas.microsoft.com/wix/2006/wi'> 
    <Product Id="#productid" 
		Language="1033" 
		Name="#productname" 
		Version="#productversion"
		Manufacturer="#manufacturer"
		UpgradeCode="#upgradecode">
        <Package Description="#productname Installer" 
		InstallPrivileges="elevated" Comments="#productshortname Installer" 
		InstallerVersion="200" Compressed="yes">
	</Package>
        <Upgrade Id="#upgradeid">
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
        <Media Id="1" Cabinet="#cabfilename" EmbedCab="yes"></Media>
        #licensecmd
        <Directory Id="TARGETDIR" Name="SourceDir">
		#installdir
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

$powershell_user_module_installdir = @"
<Directory Id="PersonalFolder" Name="PersonalFolder">
    <Directory Id="WindowsPowerShell" Name="WindowsPowerShell">
        <Directory Id="INSTALLDIR" Name="Modules">
            <Directory Id="#productshortname" Name="#modulefoldername">
            </Directory>
        </Directory>
    </Directory>
</Directory>
"@

$program_files_installdir =@"
<Directory Id="ProgramFilesFolder">
        <Directory Id="INSTALLDIR" Name="#progfilessubfolder">
            <Directory Id="#productshortname" Name="#productshortname">
            </Directory>
        </Directory>
</Directory>

"@

		#this has to be done first
		if ($InstallLocationType -eq "PowerShellUserModule")
		{
			$modules_xml = $modules_xml -replace "#installdir", $powershell_user_module_installdir
		}
		elseif( ($InstallLocationType -eq "Default") -or ($InstallLocationType -eq "ProgramFiles"))
		{
			if ( ($ProgramFilesSubFolder -eq $null) -or ($ProgramFilesSubFolder -eq ""))
			{
				Write-Host $ProgramFilesSubFolder is null
				Break
			}
			$modules_xml = $modules_xml -replace "#installdir", $program_files_installdir
		}
		else
		{
			Write-Host Unsupported InstallType
			Break
		}


		$modules_xml = $modules_xml -replace "#productid", $ProductID
	    $modules_xml = $modules_xml -replace "#productname", $ProductNameLong
	    $modules_xml = $modules_xml -replace "#productversion", $ProductVersion
	    $modules_xml = $modules_xml -replace "#manufacturer", $Manufacturer
	    $modules_xml = $modules_xml -replace "#upgradecode", $UpgradeCode
	    $modules_xml = $modules_xml -replace "#productshortname", $ProductNameShort
	    $modules_xml = $modules_xml -replace "#upgradeid", $UpgradeID
	    $modules_xml = $modules_xml -replace "#cabfilename", $cabfilename
	    $modules_xml = $modules_xml -replace "#licensecmd", $licensecmd
	    $modules_xml = $modules_xml -replace "#helplink", $HelpLink
	    $modules_xml = $modules_xml -replace "#aboutlink", $AboutLink
	    $modules_xml = $modules_xml -replace "#licensecmd", $licensecmd
		$modules_xml = $modules_xml -replace "#progfilessubfolder", $ProgramFilesSubFolder
		$modules_xml = $modules_xml -replace "#modulefoldername", $ModuleFolderName

	    $modules_xml = [xml] $modules_xml
	    $modules_xml.Save( $modules_wxs )
		
        # ----------------------------------------
        # PRODUCE THE WXS FILE
        Write-Host Writing the modules WXS file $modules_wxs
        $modules_xml.Save( $modules_wxs )
        AssertFileWasProduced( $modules_wxs )

        Write-Host Using HEAT.EXE to create the product WXS file $product_wxs 
        &$heatexe dir $InputFolder -nologo -sfrag -suid -ag -srd -dir $directoryid  -out $product_wxs -cg $ProductNameShort  -dr $ProductNameShort
        AssertFileWasProduced( $product_wxs )

        # ----------------------------------------
        # PRODUCE THE WIXOBJ FILES VIA CANDLE
        Write-Host Using CANDLE.EXE to create wixobj files
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
        if ($KeepTemporaryFolder)
		{
			# Do nothing
		}
		else
        {
			Remove-FolderIfExists -Folder $temp_folder -Description "Temp Folder" -Verbose
        }

        # These have to be manually removed because they don't go into the temp folder by default
        Remove-FileIfExists -Filename $modules_wixobj -Description "module wixobj" 
        Remove-FileIfExists -Filename  $product_wixobj -Description "product wixobj"
        Remove-FileIfExists -Filename $productpdb -Description "product pdb"

        Write-Host "----------------------------------------"
        Write-Host "Creating ZIP file"
        $zipfile = join-path $output_msi_path ($msibasename + ".zip")
        Export-ZIPFolder -InputFolder $binpath -OutputFile $zipfile -IncludeBaseDir $false


        # ----------------------------------------
        # FINAL MESSAGE
        Write-Host 
        Write-Host "----------------------------------------"
        Write-Host SUCCESS: Installer file created here $output_msi_file 

        $result = New-Object Object
        $result | Add-Member NoteProperty MSIFile $output_msi_file  
        $result | Add-Member NoteProperty ZipFile $zipfile 
        $result | Add-Member NoteProperty ProductVersion $ProductVersion                 

        return $result
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
        
        Write-Verbose $output_folder
        Write-Host "copying"

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
        {
            New-Item $output_folder -ItemType Directory
        }

        Write-Host "copying"

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


function Update-PSD1Version($Version)
{

    $src_psd1_filename = JoinResolve-Path $scriptpath $psdfilename
    $dst_psd1_filename = JoinResolve-Path $binpath $psdfilename

    if (!( Test-TextFilesAreEqual $src_psd1_filename $dst_psd1_filename))
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
            Write-Host $src_line
            $tokens = $src_line -split "="
            if ($tokens.Length -ne 2)
            {
                Write-Error "Unexpected number of tokens"
            }

            $old_version = $tokens[1].Replace("'","").Trim()
            Write-Host Old Version: $old_version
            $Version = Update-Version $old_version 2
            $new_line = "ModuleVersion = '$Version'" 
            $psd1_src[$i] = $new_line
        }
    }

    if ($Version -eq "UNKNOWN")
    {
        Write-Error "Version was never set"
    }

    Write-Host New Version: $Version

    Set-Content $src_psd1_filename $psd1_src
    Set-Content $dst_psd1_filename $psd1_src

    $Version
}

