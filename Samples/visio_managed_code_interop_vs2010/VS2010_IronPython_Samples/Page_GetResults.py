import clr 
import System 

clr.AddReference("Microsoft.Office.Interop.Visio") 
import Microsoft.Office.Interop.Visio 
IVisio = Microsoft.Office.Interop.Visio 

from Records import *
import Util

def Page_GetResults( doc ):

        pages = doc.Pages
        page = pages.Add()
        page.NameU = "PGR"

        shape = page.DrawRectangle(1, 1, 4, 3)
        shape.CellsU["Width"].Formula = "=(1.0+2.5)"
        shape.CellsU["Height"].Formula = "=(0.0+1.5)"

        # BUILD UP THE REQUEST
        flags = System.Int16(IVisio.VisGetSetArgs.visGetFloats)
        items = [

                Page_GetResults_Record( shape.ID, IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowXFormOut, IVisio.VisCellIndices.visXFormWidth, IVisio.VisUnitCodes.visNoCast),
                Page_GetResults_Record( shape.ID, IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowXFormOut, IVisio.VisCellIndices.visXFormHeight, IVisio.VisUnitCodes.visNoCast),
        ]

        # MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        SRCStream = Util.get_new_system_array(System.Int16, len(items)*4)
        unitcodes = Util.get_new_system_array(System.Object, len(items))
        for i in xrange(len(items)) :
                SRCStream[i * 4 + 0] = items[i].ShapeID
                SRCStream[i * 4 + 1] = items[i].SectionIndex
                SRCStream[i * 4 + 2] = items[i].RowIndex
                SRCStream[i * 4 + 3] = items[i].CellIndex
                unitcodes[i] = items[i].UnitCode

        # EXECUTE THE REQUEST
        results_sa = Util.get_outref_to_system_array(System.Object) 
        SRCStream_sa = Util.get_ref_to_system_array(System.Int16,SRCStream) 
        unitcodes_sa = Util.get_ref_to_system_array(System.Object,unitcodes) 
        page.GetResults(SRCStream_sa, flags, unitcodes_sa, results_sa)

        # OUTPUT BACK TO SOMETHING USEFUL 
        results = Util.get_new_system_array(System.Double,results_sa.Length)
        results_sa.CopyTo(results, 0);

        shape.Text = System.String.Format("Results={0},{1}", results[0], results[1])

