# Create a new solution via File > New Solution > Boo > Console Application 
# Select the project in the newly created solution and right-click on Project References 
# Select Add Reference 
# There are two tabs of interest: One if called GAC and the other COM 
# In the COM tab, you will see an item called Microsoft Visio 12.0 Type Library – Do not select this item. 
# Instead go to the GAC Tab and select Micropsoft.Office.Interop.Visio - Version 12.0.0.0 


namespace testvisio 


import System 
import Microsoft.Office.Interop.Visio as IVisio 

visapp = IVisio.ApplicationClass() 
doc = visapp.Documents.Add("") 
page = visapp.ActivePage 

shape = page.DrawRectangle(1, 1, 5, 4) 
shape.Text = "Hello World" 
