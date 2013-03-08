#First import the VisioPS
Import-Module D:\saveenr\code\visioautomation\VisioAutomation_2010\VisioPS\bin\Debug\VisioPS.dll
#Import-Module VisioPS

# There are a lot of Visio-related cmdlets
Get-Command -Module VisioPS

#We could connect to a Visio instance with 
#    $visapp = Connect-VisioApplication
#Now launch a new instance of Visio
New-VisioApplication

#Create a new document
New-VisioDocument

#Draw a rectangle
New-VisioRectangle 0 0 1 1



#get rid of that shape
Remove-VisioShape

#Why did that work?
#By default most of these cmdlets work on the current selection
#But you can target specific objects if needed
Get-Help Remove-VisioShape

#Let's do this the normal way
#First load the Stencil
$basic_u = Open-VisioDocument basic_u.vss

#Then Get a master from the stencil
$master = Get-VisioMaster "Rectangle" $basic_u

#Now drop the shape somewhere
$shape = New-VisioShape $master 3,3

Write-Host $shape

#Set text
Set-VisioText "Hello World"

#Let's drop a lot of shapes
Get-Help New-VisioShape

#Now drop multiple shapes
#Notice they are all selected
#Generally that will be true
$shapes = New-VisioShape $master 3,3, 5,5, 7,7

#Now you see why Set-VisioText is better
Set-VisioText "Hello World"

#New
New-VisioPage
New-VisioPage –Width 5.0 –Height 2.0 –Name “MyPage1”
Set-VisioPageLayout –Orientation Landscape
Set-VisioPageLayout –Orientation Portrait
Set-VisioPageCell -PageWidth 10 -PageHeight 5

$pages = Get-VisioPage *
$pages.Count

Remove-VisioPage

New-VisioShape $master 1,1 | Out-Null
New-VisioShape $master 3,3 | Out-Null
New-VisioShape $master 5,5 | Out-Null
New-VisioShape $master 7,7 | Out-Null

Select-VisioShape Invert
Select-VisioShape None
Select-VisioShape All
Set-VisioText -Text "A"
Set-VisioText -Text "A","B"
Set-VisioText -Text "A","B","C","D", "E" 
Set-VisioShapeCell -Width 1.0 
Get-Help Set-VisioShapeCell
Set-VisioShapeCell -CharColor "rgb(255,255,255)" -LineWeight 4pt -LinePattern 10 
Undo-Visio

Invoke-VisioAlignShape –Horizontal Left
Undo-Visio
Invoke-VisioAlignShape –Vertical Bottom
Undo-Visio

$shapes = Get-VisioShape Selected
$dc = Get-VisioMaster "Dynamic Connector" $basic_u


New-VisioConnection -From $shapes[0] -To $shapes[1] -Master $dc
Undo-Visio
New-VisioConnection -From $shapes[0],$shapes[1] -To $shapes[1],$shapes[2] -Master $dc

Set-VisioCustomProperty -Name "Prop1" -Value "Prop2"


Select-VisioShape All
Remove-VisioShape

$grid= New-VisioGridLayout -Master $master -Columns 4 -Rows 6 -CellWidth 1.0 -CellHeight 0.5 
Invoke-VisioDraw -GridLayout $grid

Undo-Visio

#To draw an Organizational chart. Just create an XML as shown below:
notepad orgchart1.xml

#Then load the XML:
$orgchart= Import-VisioOrgChart -Filename orgchart1.xml

#Then the Invoke-VisioDraw cmdlet to render it.
Invoke-VisioDraw -OrgChart $orgchart
