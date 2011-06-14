import clr 
import System 

clr.AddReference("Microsoft.Office.Interop.Visio") 
import Microsoft.Office.Interop.Visio 
IVisio = Microsoft.Office.Interop.Visio 

from Records import *
from Util import *

import Shape_GetFormulas
import Shape_GetResults
import Shape_SetFormulas
import Shape_SetResults
import Page_GetFormulas
import Page_GetResults
import Page_SetFormulas
import Page_SetResults
        
def test() :
    visapp = IVisio.ApplicationClass() 
    doc = visapp.Documents.Add("") 

    # shape
    Shape_GetFormulas.Shape_GetFormulas(doc)
    Shape_GetResults.Shape_GetResults(doc)          
    Shape_SetFormulas.Shape_SetFormulas(doc)
    Shape_SetResults.Shape_SetResults(doc)
    
    # page
    Page_GetFormulas.Page_GetFormulas(doc)
    Page_GetResults.Page_GetResults(doc)
    Page_SetFormulas.Page_SetFormulas(doc)
    Page_SetResults.Page_SetResults(doc)
        
if __name__ == "__main__" :
        test()
        
