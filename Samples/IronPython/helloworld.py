import sys
import clr
import System

# Load Visio Primary Interop Assembly
clr.AddReference("Microsoft.Office.Interop.Visio")
import Microsoft.Office.Interop.Visio
IVisio = Microsoft.Office.Interop.Visio


# Create a new instance of the application
visapp = IVisio.ApplicationClass()


# On a new doc, get the first page, and draw a rectangle
doc = visapp.Documents.Add("")
page = visapp.ActivePage
shape = page.DrawRectangle(1, 1, 5, 4)
shape.Text = "Hello World"
