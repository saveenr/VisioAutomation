#First import the VisioPS
Import-Module VisioPS

#Now launch a new instance of Visio
New-VisioApplication

#Create a new document
New-VisioDocument

#Draw a rectangle
New-VisioRectangle 0 0 1 1

#Set ext
Set-VisioText "Hello World"
