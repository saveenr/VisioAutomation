$visio_psd1 = join-path $PSScriptRoot ".\bin\Debug\Visio.psd1"
Import-Module $visio_psd1 




$shape_cells = New-VisioShapeCells
$shape_cells | select *

break

# Get a new document ready
New-VisioApplication
New-VisioDocument



$shape_cells1 = New-VisioShapeCells
$shape_cells2 = New-VisioShapeCells
$shape_cells3 = New-VisioShapeCells

$shape_cells = @( $shape_cells1, $shape_cells2, $shape_cells3)

$shape_cells1.XFormWidth = 2
$shape_cells1.FillForeground = "rgb(255,255,0)"
$shape_cells2.XFormHeight = 4
$shape_cells2.FillForeground = "rgb(0,0,255)"
$shape_cells3.FillForeground = "rgb(255,0,0)"
$shape_cells3.LineWeight = "5 pt"


$basic_u = Open-VisioDocument "basic_u.vss"
$masters = Get-VisioMaster -Name "Rectangle","Triangle","Circle" $basic_u
$positions = @(4,5,6,1,0,0)
$shapes = New-VisioShape $masters $positions -Cells $shape_cells

