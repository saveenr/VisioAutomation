Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"

Import-Module .\CodePackage.psm1

$mypath = $MyInvocation.MyCommand.path
$visioautomation_path = Resolve-Path ( Join-Path $MyInvocation.MyCommand.path "..\..\.." )
$bindebug_path = Resolve-Path( Join-Path $visioautomation_path  "visioautomation_2010\VisioPS\bin\Debug" )


Install-PSModuleFromFolder $bindebug_path "VisioPS" -Verbose