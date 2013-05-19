cls
CD \

$script_path = $myinvocation.mycommand.path
$script_folder = Split-Path $script_path -Parent

#Write-Host $script_folder
$bin_debug = Join-Path $script_folder "../bin/debug"
#Write-Host $bin_debug

$visiops_dll = Join-Path $bin_debug "VisioPS.dll"
#Write-Host $visiops_dll 


$types_file = Join-Path $bin_debug "VisioPS.Types.ps1xml"

Import-Module $visiops_dll 

Update-TypeData $types_file

New-VisioApplication
New-VisioDocument

$doc = New-VisioDocument
$basic_stencil = Open-VisioDocument "basic_u.vss"

$rectangle_master = Get-VisioMaster -Master "Rectangle" -Stencil $basic_stencil


$shapes = New-VisioShape -Masters $rectangle_master -Points 2.0,2.0,4.0,4.0,6.0,6.0

Set-VisioShapeCell -FillForegnd "rgb(255,0,0)" -Width "1" -CharSize "20pt" -Shapes $shapes[0]
Set-VisioShapeCell -FillForegnd "rgb(255,128,50)" -Width "2" -CharSize "30pt" -Shapes $shapes[1]
Set-VisioShapeCell -FillForegnd "rgb(255,200,50)" -Width "3" -CharSize "40pt" -Shapes $shapes[2]

Set-VisioText -Shapes $shape -Text "A","B","C"