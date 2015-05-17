# Custom Properties (Shape Data)

	Set-VisioCustomProperty "foo" "BAR"
# Get Custom Properties

	$props = Get-VisioCustomProperty

# Here is a more complete example.

	$doc = New-VisioDocument
	$stencil_net = Open-VisioDocument "Basic Network Diagram.vst"
	$stencil_comp = Open-VisioDocument "Computers and Monitors.vss"
	
	$pc_master = Get-VisioMaster -Master "PC" -Stencil $stencil_comp 
	
	$shapes = New-VisioShape -Masters $pc_master -Points 2.2,6.8
	$shape1 = $shapes[0]
	
	Select-VisioShape -Shapes $shape1
	$shape1.Text = "Some Text..."
	
	Set-VisioCustomProperty -Name "prop1" -Value "val1"
	Set-VisioCustomProperty -Name "prop2" -Value "val2"
	
	$shapedata = Get-VisioCustomProperty
	
	$props_for_shape1 = $shapedata[ $shape1]
	
	foreach ($propname in $props_for_shape1.Keys)
	{
	    $custompropcells = $props_for_shape1[ $propname ]
	    Write-Host $propname = $custompropcells.Value.Formula
	}

# Delete a Custom Property

	Remove-VisioCustomProperty "foo"

