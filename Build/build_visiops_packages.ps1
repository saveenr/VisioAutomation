Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"

#Remove-Module CodePackage
Import-Module .\CodePackage.psm1

$verstring = "1.1.8"
$mypath = $MyInvocation.MyCommand.path
$visioautomation_path = Resolve-Path ( Join-Path $MyInvocation.MyCommand.path "..\.." )
$bindebug_path = Resolve-Path( Join-Path $visioautomation_path  "visioautomation_2010\VisioPS\bin\Debug" )
$wixbin_path = Resolve-Path( Join-Path $visioautomation_path  "Build\wix36-binaries" )

$mydocs = join-Path $env:USERPROFILE Documents
$output_folder = Join-Path $mydocs "Visio Powershell Distribution"
$zipfile = Join-Path $output_folder ( "VisioPS_" + $verstring + ".zip")

if (!(Test-Path $output_folder)) {
    New-Item -Path $output_folder -ItemType Directory
}

Export-PowerShellModuleInstaller `
    -InputFolder $bindebug_path `
    -OutputFolder $output_folder `
    -WIXBinFolder $wixbin_path `
	-InstallType "PowerShellUserModule" `
    -ProductNameLong "Visio Powershell Module" `
    -ProductNameShort "VisioPS" `
    -ProductVersion $verstring  `
    -Manufacturer "Saveen Reddy" `
    -HelpLink "http://visioautomation.codeplex.com" `
    -AboutLink  "http://visioautomation.codeplex.com" `
    -ProductID "4A2B528A-93E5-431D-97BB-79767C7677C5" `
    -UpgradeCode  "EE659AB6-BE76-426E-B971-35DF3907F9D4" `
    -UpgradeID  "F14DC5AF-1234-498A-9646-AA27E03957AA" `
	-ProgramFilesSubFolder $null `
    -KeepTemporaryFolder $false 


Export-ZIPFolder -InputFolder $bindebug_path -OutputFile $zipfile -IncludeBaseDir $false -Overwrite