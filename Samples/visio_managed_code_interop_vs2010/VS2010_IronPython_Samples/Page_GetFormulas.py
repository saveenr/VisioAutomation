import clr 
import System 

clr.AddReference("Microsoft.Office.Interop.Visio") 
import Microsoft.Office.Interop.Visio 
IVisio = Microsoft.Office.Interop.Visio 

from Records import *
import Util

def Page_GetFormulas( doc ):

        pages = doc.Pages
        page = pages.Add()
        page.NameU = "PGF"
        
        shape = page.DrawRectangle(1, 1, 4, 3)
        shape.CellsU["Width"].Formula = "=(1.0+2.5)"
        shape.CellsU["Height"].Formula = "=(0.0+1.5)"

        # BUILD UP THE REQUEST
        items = [

                Page_GetFormulas_Record( shape.ID, IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowXFormOut, IVisio.VisCellIndices.visXFormWidth),
                Page_GetFormulas_Record( shape.ID, IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowXFormOut, IVisio.VisCellIndices.visXFormHeight)
        ]

        # MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        SRCStream = Util.get_new_system_array(System.Int16, len(items)*4)
        for i in xrange(len(items)) :
                SRCStream[i * 4 + 0] = items[i].ShapeID
                SRCStream[i * 4 + 1] = items[i].SectionIndex
                SRCStream[i * 4 + 2] = items[i].RowIndex
                SRCStream[i * 4 + 3] = items[i].CellIndex

        # EXECUTE THE REQUEST
        formulas_sa = Util.get_outref_to_system_array(System.Object) 
        SRCStream_sa = Util.get_ref_to_system_array(System.Int16,SRCStream)  
        page.GetFormulasU(SRCStream_sa, formulas_sa)

        # OUTPUT BACK TO SOMETHING USEFUL 
        formulas = Util.get_new_system_array(System.String,formulas_sa.Length)
        formulas_sa.CopyTo(formulas, 0);

        shape.Text = System.String.Format("Formulas={0},{1}", formulas[0], formulas[1])

