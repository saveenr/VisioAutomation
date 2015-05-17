# Connectors

# Connecting Shapes

Select two shapes and then use Connect-Shape

	$basic_stencil = Open-VisioDocument "basic_u.vss"
	$rectangle_master = Get-VisioMaster -Master "Rectangle" -Stencil $basic_stencil
	$dyncon_master = Get-VisioMaster -Master "Dynamic Connector" -Stencil $basic_stencil 
	$shapes = New-VisioShape -Masters $rectangle_master -Points 2.0,2.0,4.0,4.0
	New-VisioConnection -From $shape[0] -To $shape[1] -Master $dyncon_master

