# Using the ShapeSheet

There are four useful cmdlets for examining and editing the ShapeSheets. Two cmdlets work on Shape objects and two cmdlets work on Page objects.
* Get-VisioShapeCell
* Set-VisioShapeCell
* Get-VisioPageCell
* Set-VisioPageCell

# Retrieving Cell Values


	Get-VisioShapeCell -Cell FillForegnd

Output:
	FillForegnd 
	-----------
	RGB(255,200,50)
	RGB(255,128,50)
	RGB(255,0,0)

Multiple values can be retrieved


	Get-VisioShapeCell -Cell FillForegnd,FillPattern


	FillForegnd     FillPattern
	-----------     -----------
	RGB(255,200,50) 1
	RGB(255,128,50) 1   
	RGB(255,0,0)    1

With Get-VisioShapeCell you can also specify which shapes to look at via the -Shapes parameter, otherwise the active selection is used.

Below are examples of how to retrieve cell values with Get-VisioPageCell

	Get-VisioPageCell -Cell PageWidth,PageHeight

Will return

	PageWidth PageHeight
	--------- ----------    
	8.5 in    11 in

The Get-VIsioPageCell cmdlet always only works with the active page.
Both Get-VIsioPageCell and Get-VIsioPageCell by default return formulas for those cells. 

This is clear if you get the LocPinX amd LocPinY cells.

	Get-VisioShapeCell -Cell LocPinX,LocPinY

	LocPinX    LocPinY
	-------    -------
	Width*0.5  Height*0.5
	Width*0.5  Height*0.5    
	Width*0.5  Height*0.5

If want the results, use the -GetResults parameter.


	Get-VisioShapeCell LocPinX,LocPinY -GetResults

	LocPinX     LocPinY	
	-------     -------
	1.5000 in.  0.5000 in.
	1.0000 in.  0.5000 in.
	0.5000 in.  0.5000 in.

By default the results are returned as strings. However, this can be controlled with the -ResultType parameter which supports these values: Double, Integer, String

	Get-VisioShapeCell LocPinX, LocPinY -GetResults -ResultType Double

	LocPinX LocPinY
	------- ------- 
	1.5     0.5
	1       0.5
	0.5     0.5
	

Uses of the -Cell Parameter
The -Cell parameter allows wildcards to retrieve multiple cells

	Get-VisioShapeCell -Cell *

You can even use multiple Wildcard parameters

	Get-VisioShapeCell -Cell Fill*,Lock*

You can mix the cell specific switches with the wildcards

	Get-VisioShapeCell -PinX -Cell Fill*,Lock*

Setting Cell Values
To set the value of a ShapeSheet cell use Set-VisioShapeCell. Each cell is a separate parameter and represents the formula for that cell

	Set-VisioShapeCell -FillForegnd "rgb(255,0,0)"

Multiple cells can be set at the same time
	Set-VisioShapeCell -FillForegnd "rgb(255,0,0)" -Width "1"
By default the cells will be set on the selected shapes, but by using the -Shapes parameter you can identify the specific shapes to use without regard to selection

	Set-VisioShapeCell -FillForegnd "rgb(255,0,0)" -Width "1" -Shapes $shapes

Using Set-VisioPageCell to set page properties. In the example below, the active page is resized.

	Set-VisioPageCell -PageWidth 3 -PageHeight 4 -PageWidth "3" -CharSize "40pt" -Shapes $shapes[2]
