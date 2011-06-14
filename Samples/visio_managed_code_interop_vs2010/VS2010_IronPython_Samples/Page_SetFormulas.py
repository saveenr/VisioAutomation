import clr 
import System 

clr.AddReference("Microsoft.Office.Interop.Visio") 
import Microsoft.Office.Interop.Visio 
IVisio = Microsoft.Office.Interop.Visio 

from Records import *
import Util

def Page_SetFormulas( doc ):

        pages = doc.Pages
        page = pages.Add()
        page.NameU = "PSF"
        
        shape = page.DrawRectangle(1, 1, 4, 3)
        shape.CellsU["Width"].Formula = "=(1.0+2.5)"
        shape.CellsU["Height"].Formula = "=(0.0+1.5)"

        # BUILD UP THE REQUEST
        flags = System.Int16(IVisio.VisGetSetArgs.visSetBlastGuards | IVisio.VisGetSetArgs.visSetUniversalSyntax)        
        items = [
                Page_SetFormulas_Record( shape.ID , IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowXFormOut, IVisio.VisCellIndices.visXFormWidth, "1.3"),
                Page_SetFormulas_Record( shape.ID , IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowXFormOut, IVisio.VisCellIndices.visXFormHeight, "7.71")
        ]

        # MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        SRCStream = Util.get_new_system_array(System.Int16, len(items)*4)
        formulas = Util.get_new_system_array(System.Object, len(items))
        for i in xrange(len(items)) :
                SRCStream[i * 4 + 0] = items[i].ShapeID
                SRCStream[i * 4 + 1] = items[i].SectionIndex
                SRCStream[i * 4 + 2] = items[i].RowIndex
                SRCStream[i * 4 + 3] = items[i].CellIndex
                formulas[i] = items[i].Formula

        # EXECUTE THE REQUEST
        formulas_sa = Util.get_ref_to_system_array(System.Object,formulas) 
        SRCStream_sa = Util.get_ref_to_system_array(System.Int16,SRCStream)  
        count = page.SetFormulas(SRCStream_sa, formulas_sa, flags)

        shape.Text = System.String.Format("SetFormulas")


