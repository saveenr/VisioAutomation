# Selection

# Select All Shapes
	Select-VisioShape All

# Deselect All Shapes (Clear Selection)

	Select-VisioShape None

# Invert the Selection

	Select-VisioShape Invert

# Selecting Specific Shapes

	New-VisioRectangle 0 0 1 1
	$shapes += New-VisioRectangle 0 0 1 1
	$shapes += New-VisioRectangle 0 0 2 3

	Select-VisioShape $shapes

You can pass in IDs of shapes

	New-VisioRectangle 0 0 1 1
	$shapes += New-VisioRectangle 0 0 1 1
	$shapes += New-VisioRectangle 0 0 2 3
	$shapeids = $shapes | ForEach-Object{ $_.ID }
	Select-VisioShape $shapeids

# Checking if there are Selected Shapes Available

This cmdlet returns $true if there are any *selected* shapes available in the active document.

	If (Test-VisioSelectedShapes)
	{
	    # do something  
	}

# Getting Selected Shapes

	$shapes = Get-VisioShape

# Get All Shapes on Page regardless of Selection

	$shapes = Get-VisioShape *

# Get All Shapes on Page by Name

	$shapes = Get-VisioShape "Shapename"

NOTE: Wildcards are NOT supported

# Get Selected Shapes included Shapes inside Groups

	$shapes = Get-VisioShape -Recursive
	

