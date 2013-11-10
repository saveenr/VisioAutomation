Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"

cls




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

function Assert-Equals( $expected, $v ,$msg=$null)
{
    if ($v -ne $expected)
    {
	        Write-Host "ERROR: Assert Failed Expected" $expected "actually got" $v " :" $msg -ForegroundColor Red
    }
}

Assert-VisioPSIsInstalled
Load-VisioPSModule
Create-VisioApplication


function Assert-PageShapeCount( $desired )
{
    $shapes = Get-VisioShape -Flags Page
    $actual = $shapes.Count

    if ($actual -ne $desired)
    {
        Write-Host ERROR: Expected $desired shapes on page, got $actual -ForegroundColor Gray
        exit
    }
}

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


$Test1 = 
{ 
    Write-Host "Test: Start Test Placeholder"
    New-VisioDocument
    #Throw [system.IndexOutOfRangeException]  
} 

$Test2 = 
{ 
    New-VisioPage
    Write-Host "Test: Get-Edges"
    $r0 = New-VisioRectangle 0 0 1 1
    $r1 = New-VisioRectangle 3 3 4 4 
    $r2 = New-VisioRectangle 6 6 7 7

    New-VisioConnection -From $r0 -To $r1 -Verbose 
    New-VisioConnection -From $r1 -To $r2 -Verbose 

    Assert-PageShapeCount 5

    $edges = Get-VisioEdge

    Assert-Equals 2 $edges.Count 
    Assert-Equals 1 $edges[0].FromShapeID 
    Assert-Equals 2 $edges[0].ToShapeID   
    Assert-Equals 2 $edges[1].FromShapeID 
    Assert-Equals 3 $edges[1].ToShapeID   

    Remove-VisioPage 
} 

$Test3 = 
{ 
    $doc = Get-VisioDocument -ActiveDocument
    $pages = $doc.Pages
    $oldpagecount = $pages.Count

    New-VisioPage
    Assert-Equals ($oldpagecount+1) $pages.Count

    Write-Host "Test: Page Duplication"
    $r0 = New-VisioRectangle 0 0 1 1
    $r1 = New-VisioRectangle 3 3 4 4 
    $r2 = New-VisioRectangle 6 6 7 7
	Set-VisioText  -Shapes $r0,$r1,$r2 "Shape0", "Shape1", "Shape2"
    Assert-PageShapeCount 3

    Invoke-VisioDuplicatePage
    Assert-Equals ($oldpagecount+2) $pages.Count
    Remove-VisioPage 

    Assert-Equals ($oldpagecount+1) $pages.Count
    Remove-VisioPage 
    Assert-Equals ($oldpagecount) $pages.Count
} 

$Test4 = 
{ 
    New-VisioPage

    Write-Host "Test: Page Duplication"
    $r0 = New-VisioRectangle 0 0 1 1
    $r1 = New-VisioRectangle 3 3 4 4 
    $r2 = New-VisioRectangle 6 6 7 7
	Set-VisioText  -Shapes $r0,$r1,$r2 "Shape0", "Shape1", "Shape2"
    Assert-PageShapeCount 3

    $p0 = @{ PString = "HelloWorld" }

   
    Select-VisioShape -Operation None
    Select-VisioShape $r0 
    Set-VisioCustomProperty $p0

    $now = Get-Date
    $p1 = @{ PDate = $now ; PInt=7 ; PDouble=3.14 }

    Select-VisioShape -Operation None
    Select-VisioShape $r1
    Set-VisioCustomProperty $p1

    Select-VisioShape -Operation All
    $ap0 = Get-VisioCustomProperty
    Remove-VisioPage 
} 

$Test5 = 
{ 
    New-VisioPage

    Write-Host "Test: Get Specific  Shapes"
    $r0 = New-VisioRectangle 0 0 1 1
    $r1 = New-VisioRectangle 3 3 4 4 
    $r2 = New-VisioRectangle 6 6 7 7

	Select-VisioShape -Operation None
	Set-VisioText  -Shapes $r0,$r1,$r2 "Shape0","Shape1","Shape2"

	$x = Get-VisioShape 1,2,3 #by ID
	$y = Get-VisioShape "Sheet.1","Sheet.2","Sheet.3"  #by name

	Assert-Equals "Shape0" $x[0].Text
	Assert-Equals "Shape1" $x[1].Text
	Assert-Equals "Shape2" $x[2].Text
	Assert-Equals "Shape0" $y[0].Text
	Assert-Equals "Shape1" $y[1].Text
	Assert-Equals "Shape2" $y[2].Text

	Remove-VisioPage 
} 


Run-Test $Test1
#Run-Test $Test2
#Run-Test $Test3
#Run-Test $Test4
Run-Test $Test5

Write-Host
Write-Host ------ -ForegroundColor Yellow
Write-Host Ending -ForegroundColor Yellow
Write-Host ------ -ForegroundColor Yellow






Close-VisioApplication -Force


Assert-True ($null -eq (Get-VisioApplication)) "Visio Application is still referenced after force close"
