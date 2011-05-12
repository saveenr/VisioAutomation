import clr 
import System 

clr.AddReference("Microsoft.Office.Interop.Visio") 
import Microsoft.Office.Interop.Visio 

IVisio = Microsoft.Office.Interop.Visio 

visapp = IVisio.ApplicationClass() 
doc = visapp.Documents.Add("") 
page = visapp.ActivePage 

shape = page.DrawRectangle(1, 1, 5, 4) 
shape.Text = "Hello World" 
