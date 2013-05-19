Param($demotype)
Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"

$Host.UI.RawUI.WindowTitle = "PowerShell:VisioPS"
CD \
cls

$script_path = $myinvocation.mycommand.path
$script_folder = Split-Path $script_path -Parent


if ($demotype -eq "bindebug")
{
    $bin_debug = Join-Path $script_folder "../bin/debug"
    $visiops_psd1 = Join-Path $bin_debug "VisioPS.psd1"
    Import-Module $visiops_psd1
}
else
{
    Import-Module VisioPS
}


$visiopsmodule = Get-Module VisioPS

New-VisioApplication

$doc = New-VisioDocument
$basic_stencil = Open-VisioDocument "basic_u.vss"

$rectangle_master = Get-VisioMaster -Master "Rectangle" -Stencil $basic_stencil
$dyncon_master = Get-VisioMaster -Master "Dynamic Connector" -Stencil $basic_stencil


$shapes = New-VisioShape -Masters $rectangle_master -Points 2.0,2.0,4.0,4.0,6.0,6.0

Set-VisioShapeCell -FillForegnd "rgb(255,0,0)" -Width "1" -CharSize "20pt" -Shapes $shapes[0]
Set-VisioShapeCell -FillForegnd "rgb(255,128,50)" -Width "2" -CharSize "30pt" -Shapes $shapes[1]
Set-VisioShapeCell -FillForegnd "rgb(255,200,50)" -Width "3" -CharSize "40pt" -Shapes $shapes[2]

Set-VisioText -Shapes $shapes -Text "A","B","C"

$con0 = New-VisioConnection -From $shapes[0] -To $shapes[1] -Master $dyncon_master

Set-VisioShapeCell -LineWeight "3pt" -EndArrow 1 -EndArrowSize "10pt" -Shapes $con0

Write-Host Demo Type: $demotype
Write-Host Module Path: $visiopsmodule.Path
Write-Host Module Version: $visiopsmodule.Version
Write-Host
