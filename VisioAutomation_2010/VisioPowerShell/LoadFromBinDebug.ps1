Import-Module Visio

# Get a new document ready
New-VisioApplication
New-VisioDocument

$basic_u = Open-VisioDocument "basic_u.vss"
$rect_master = Get-VisioMaster "Rectangle" $basic_u
$triangle_master = Get-VisioMaster "Triangle" $basic_u
$circle_master = Get-VisioMaster "Circle" $basic_u
$masters = @( $rect_master, $triangle_master , $circle_master )
$positions = @(4,5,6,1,0,0)
$shapes = New-VisioShape $masters $positions