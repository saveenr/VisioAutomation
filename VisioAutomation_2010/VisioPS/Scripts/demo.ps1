#First import the VisioPS
Import-Module VisioPS

# There are a lot of Visio-related cmdlets
Get-Command -Module VisioPS | Out-GridView

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
$basic_u = Open-VisioStencil basic_u.vss

#Then Get a master from the stencil
$master = Get-VisioMaster "Rectangle" $basic_u

#Now drop the shape somewhere
$shape = New-VisioShape $master 3,3

#Set text
$shape.Text = "Hello"

#Set text (a better way)
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

$pages = Get-VisioPage *
$pages.Count

Remove-VisioPage
Remove-VisioPage $pages

#Close Notice it works on the current document
Close-VisioDocument
