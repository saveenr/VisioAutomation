import clr 
import System 

clr.AddReference("Microsoft.Office.Interop.Visio") 
import Microsoft.Office.Interop.Visio 
IVisio = Microsoft.Office.Interop.Visio 

from Records import *
import Util


def Page_SetResults( doc ):

        pages = doc.Pages
        page = pages.Add()
        page.NameU = "PSR"
        
        shape = page.DrawRectangle(1, 1, 4, 3)
        shape.CellsU["Width"].Formula = "=(1.0+2.5)"
        shape.CellsU["Height"].Formula = "=(0.0+1.5)"

        # BUILD UP THE REQUEST
        flags = System.Int16(0)        
        items = [
                Page_SetResults_Record( shape.ID, IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowXFormOut, IVisio.VisCellIndices.visXFormWidth, 8.0, IVisio.VisUnitCodes.visNoCast),                
                Page_SetResults_Record( shape.ID, IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowXFormOut, IVisio.VisCellIndices.visXFormHeight, 1.0, IVisio.VisUnitCodes.visNoCast)
        ]

        # MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        SRCStream = Util.get_new_system_array(System.Int16, len(items)*4)
        results = Util.get_new_system_array(System.Object, len(items))
        unitcodes = Util.get_new_system_array(System.Object, len(items))
        for i in xrange(len(items)) :
                SRCStream[i * 4 + 0] = items[i].ShapeID
                SRCStream[i * 4 + 1] = items[i].SectionIndex
                SRCStream[i * 4 + 2] = items[i].RowIndex
                SRCStream[i * 4 + 3] = items[i].CellIndex
                results[i] = items[i].Result
                unitcodes[i] = items[i].UnitCode

        # EXECUTE THE REQUEST
        results_sa = Util.get_ref_to_system_array(System.Object,results)
        unitcodes_sa = Util.get_ref_to_system_array(System.Object,unitcodes)
        SRCStream_sa = Util.get_ref_to_system_array(System.Int16,SRCStream)  
        count = page.SetResults(SRCStream_sa, unitcodes_sa, results_sa, flags)

        shape.Text = System.String.Format("SetResults")

