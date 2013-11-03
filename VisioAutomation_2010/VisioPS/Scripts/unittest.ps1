Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"

cls


function Run-Test( $sb )
{
    Write-Host ---------------------------------------- -ForegroundColor Cyan
    $test_passed = $false
    try
    {
            &$sb
            $test_passed = $true
    }
    finally
    {
        if ($test_passed)
        {
            Write-Host Passed -ForegroundColor Green
        }
        else
        {
            Write-Host Failed  -ForegroundColor Red
        }

    }
    Write-Host ---------------------------------------- -ForegroundColor Cyan
}


function prompt
{
    "VisiosPSUnitTest>"
}

$script_path = $myinvocation.mycommand.path
$script_folder = Split-Path $script_path -Parent

$bin_debug = Join-Path $script_folder "../bin/debug"
$docfolder =  "$home\documents"
$wps =  Join-Path $docfolder "WindowsPowerShell"
$modules =  Join-Path $wps "Modules"
$visiopsfldr=  Join-Path $modules "VisioPS"


Write-Host "-----------------" -ForegroundColor Gray
Write-Host "VisioPS Unit Test" -ForegroundColor Gray
Write-Host "-----------------" -ForegroundColor Gray
Write-Host
Write-Host "VisioPS Location" $visiopsfldr -ForegroundColor Gray

function Assert-PathExists($path)
{
    if (!(test-path $path))
    {
        Write-Host "ERROR: Path does not exist " + $path -ForegroundColor Gray
        exit
    }
}

$visiopsfiles = @()
$visiopsfiles += "VisioPS.dll"
$visiopsfiles += "VisioPS.psd1"
$visiopsfiles += "VisioPS.Types.ps1xml"
$visiopsfiles += "VisioAutomation.dll"
$visiopsfiles += "VisioAutomation.Scripting.dll"

function Assert-VisioPSIsInstalled
{
    Write-Host Checking VisioPS is installed -ForegroundColor Gray
    Assert-PathExists $visiopsfldr
    foreach ($file in $visiopsfiles)
    {
        Assert-PathExists (Join-Path $visiopsfldr $file)
    }
    Write-Host Installation OK -ForegroundColor Gray
}

function Load-VisioPSModule
{
    $visiopsmodule = Get-Module VisioPS

    if ($visiopsmodule -ne $null)
    {
        Write-Host "WARNING: VisioPS Module is already loaded" -ForegroundColor Yellow
    }

    Import-Module VisioPS
    $visiopsmodule = Get-Module VisioPS

    if ($visiopsmodule -eq $null)
    {
        Write-Host "ERROR: Failed to Load VisioPS" -ForegroundColor Red
        Exit
    }
}

function Create-VisioApplication
{
	#Initially there should no Visio Application installed

	if ( (Get-VisioApplication) -ne $null)
	{
	    Write-Host "WARNING: Bound Visio Application will be terminated" -ForegroundColor Yellow
	    Close-VisioApplication -Force
	}

	New-VisioApplication

	$visapp = Get-VisioApplication

	if (!(Test-VisioApplication))
	{
	        Write-Host "ERROR: Application isntance is invalid" -ForegroundColor Red
	}


	$docs = Get-VisioDocument *
	if ($docs.Count -ne 0)
	{
	        Write-Host "ERROR: Should have zero docs open" -ForegroundColor Red
	}
	Write-Host $docs

	Write-Host $visapp
}

function Assert-True( $v ,$msg=$null)
{
    if ($v -eq $false)
    {
	        Write-Host "ERROR: Assert Failed" $msg -ForegroundColor Red
    }
}
Assert-VisioPSIsInstalled
Load-VisioPSModule
Create-VisioApplication



$Test1 = 
{ 
    Write-Host "Test: Start Test Placeholder"
    New-VisioDocument
    #Throw [system.IndexOutOfRangeException]  
} 

$Test2 = 
{ 
    $r0 = New-VisioRectangle 0 0 1 1
    $r1 = New-VisioRectangle 4 5 5 5

    Write-Host "Test: Start Test Placeholder"
} 

Run-Test $Test1
Run-Test $Test2

Write-Host
Write-Host ------ -ForegroundColor Yellow
Write-Host Ending -ForegroundColor Yellow
Write-Host ------ -ForegroundColor Yellow





Close-VisioApplication -Force


Assert-True ($null -eq (Get-VisioApplication)) "Visio Application is still referenced after force close"
