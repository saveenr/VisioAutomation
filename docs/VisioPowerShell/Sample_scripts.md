
# Draw an grid on a 8.5x11 drawing

	New-VisioApplication
	New-VisioDocument
	Set-VisioPageLayout -width 8.5 -height 11
	
	# Draw the vertical lines
	for ($x=0; $x -le 8.5; $x++) { New-VisioLine $x 0 $x 11 }
	
	# Draw the horzontal lines
	for ($y=0; $y -le 11; $y++) { New-VisioLine 0 $y 8.5 $y }


# Draw all Fill Patterns
	Import-Module Visio
	New-VisioApplication
	New-VisioDocument
	New-VisioPage
	
	$numcols = 6
	$cellwidth = 0.5
	$cellsep = 1.0
	
	$d = $cellwidth + $cellsep
	for ($i=0;$i -le 40;$i++) 
	{
	    $x = $i % $numcols 
	    $y = [math]::floor($i / $numcols )
	    $left = $x*$d
	    $bottom = $y*$d
	    $right = $left + $cellwidth
	    $top = $bottom + $cellwidth
	    $s1 = New-VisioRectangle $left $bottom $right $top
	    Set-VisioShapeCell -FillForegnd  "rgb(0,128,195)" -FillBkgnd "rgb(255,255,255)" -FillPattern $i
	    $s2 = New-VisioRectangle ($left-$cellwidth) $bottom ($right-$cellwidth) $top
	    Set-VisioText $i
	}
	
	Invoke-VisioResizePageToFitContents -BorderWidth 1.0 -BorderHeight 1.0
	
