# Masters and Stencils

# Get a Master in the Active Document

	$master = Get-VisioMaster "Foo"

# Get all the Masters in the Active Document

	$masters = Get-VisioMaster
# Get a Master in a specific Stencil Document

	$master = Get-VisioMaster "Rectangle" "Basic_u.vss"
	Note: That stencil must be loaded

# Load a Stencil Document
	$stencil = Open-VisioDocument "basic_u.vss"

# Dropping a Master onto a Page
Drops a rectangle at position (5.0,3.5)

	$master = Get-VisioMaster "Rectangle" "Basic_u.vss"
	$point = 5.0,3.5
	New-VisioShape $master $point


# Dropping Multiple Masters onto a Page
Drops a rectangle at position (1.0,2.5) and a triangle at position (5.0,3.5)

	$m_rect = Get-VisioMaster "Rectangle" "Basic_u.vss" 
	$m_triangle = Get-VisioMaster "Triangle" "Basic_u.vss"
	$masters = ($m_rect, $m_triangle)
	$points = (1.0, 2.0, 5.0, 3.5)
	New-VisioShape $masters $points

New-VisioShape returns a list of integers - these are the shape ids of the shapes that were created as a result of the drop operation. 


