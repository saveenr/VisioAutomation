import clr 
import System 

clr.AddReference("Microsoft.Office.Interop.Visio") 
import Microsoft.Office.Interop.Visio 
IVisio = Microsoft.Office.Interop.Visio 

from Records import *
import Util

def Shape_SetFormulas( doc ):

        pages = doc.Pages
        page = pages.Add()
        page.NameU = "SSF"
        
        shape = page.DrawRectangle(1, 1, 4, 3)
        shape.CellsU["Width"].Formula = "=(1.0+2.5)"
        shape.CellsU["Height"].Formula = "=(0.0+1.5)"

        # BUILD UP THE REQUEST
        flags = System.Int16(IVisio.VisGetSetArgs.visSetBlastGuards | IVisio.VisGetSetArgs.visSetUniversalSyntax)        
        items = [

                Shape_SetFormulas_Record(IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowXFormOut, IVisio.VisCellIndices.visXFormWidth, "1.3"),
                Shape_SetFormulas_Record(IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowXFormOut, IVisio.VisCellIndices.visXFormHeight, "7.71")
        ]

        # MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        SRCStream = Util.get_new_system_array(System.Int16, len(items)*3)
        formulas = Util.get_new_system_array(System.Object, len(items))
        for i in xrange(len(items)) :
                SRCStream[i * 3 + 0] = items[i].SectionIndex
                SRCStream[i * 3 + 1] = items[i].RowIndex
                SRCStream[i * 3 + 2] = items[i].CellIndex
                formulas[i] = items[i].Formula

        # EXECUTE THE REQUEST
        formulas_sa = Util.get_ref_to_system_array(System.Object,formulas) 
        SRCStream_sa = Util.get_ref_to_system_array(System.Int16,SRCStream)  
        count = shape.SetFormulas(SRCStream_sa, formulas_sa, flags)

        shape.Text = System.String.Format("SetFormulas")
