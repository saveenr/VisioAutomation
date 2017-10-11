Import-Module .\Visio.psd1

# Get a new document ready
New-VisioApplication
New-VisioDocument

$basic_u = Open-VisioDocument "basic_u.vss"
$masters = Get-VisioMaster -Name "Rectangle","Triangle","Circle" $basic_u
$positions = @(4,5,6,1,0,0)
$shapes = New-VisioShape $masters $positions