Drawing a Grid of Shapes 

	$stencil = Open-VisioDocument "basic_u.vss" $m = Get-VisioMaster "Rectangle" $stenci
	$grid= New-VisioGridLayout -Master $m -Columns 4 -Rows 6 -CellWidth 1.0 -CellHeight 0.5 -CellHorizontalSpacing 1.0 -CellVerticalSpacing 1.5
	$grid | Out-Visio E:_Docs_Org_Charts.md Drawing an Org Chart 

