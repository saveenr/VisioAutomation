import clr 
import System 

clr.AddReference("Microsoft.Office.Interop.Visio") 
import Microsoft.Office.Interop.Visio 
IVisio = Microsoft.Office.Interop.Visio 

class Shape_GetFormulas_Record :

    def __init__( self, sec, row, cell ) :
        self.SectionIndex = System.Int16(sec)
        self.RowIndex = System.Int16(row)
        self.CellIndex = System.Int16(cell)

class Page_GetFormulas_Record :

    def __init__( self, sid, sec, row, cell ) :
        self.ShapeID = System.Int16(sid)
        self.SectionIndex = System.Int16(sec)
        self.RowIndex = System.Int16(row)
        self.CellIndex = System.Int16(cell)

class Shape_GetResults_Record :

    def __init__( self, sec, row, cell , unitcode ) :
        self.SectionIndex = System.Int16(sec)
        self.RowIndex = System.Int16(row)
        self.CellIndex = System.Int16(cell)
        self.UnitCode = System.Int16(unitcode)

class Page_GetResults_Record :

    def __init__( self, sid, sec, row, cell , unitcode ) :
        self.ShapeID = System.Int16(sid)
        self.SectionIndex = System.Int16(sec)
        self.RowIndex = System.Int16(row)
        self.CellIndex = System.Int16(cell)
        self.UnitCode = System.Int16(unitcode)


class Shape_SetFormulas_Record :

    def __init__( self, sec, row, cell , formula ) :
        self.SectionIndex = System.Int16(sec)
        self.RowIndex = System.Int16(row)
        self.CellIndex = System.Int16(cell)
        self.Formula = formula

class Page_SetFormulas_Record :

    def __init__( self, sid, sec, row, cell , formula ) :
        self.ShapeID = System.Int16(sid)
        self.SectionIndex = System.Int16(sec)
        self.RowIndex = System.Int16(row)
        self.CellIndex = System.Int16(cell)
        self.Formula = formula


class Shape_SetResults_Record :

    def __init__( self, sec, row, cell , result, unitcode ) :
        self.SectionIndex = System.Int16(sec)
        self.RowIndex = System.Int16(row)
        self.CellIndex = System.Int16(cell)
        self.Result = result
        self.UnitCode = System.Int16(unitcode)

class Page_SetResults_Record :

    def __init__( self, sid, sec, row, cell , result, unitcode ) :
        self.ShapeID = System.Int16(sid)
        self.SectionIndex = System.Int16(sec)
        self.RowIndex = System.Int16(row)
        self.CellIndex = System.Int16(cell)
        self.Result = result
        self.UnitCode = System.Int16(unitcode)
        
def Shape_GetFormulas( doc ):

        pages = doc.Pages
        page = pages.Add()
        page.NameU = "SGF"
        
        shape = page.DrawRectangle(1, 1, 4, 3)
        shape.CellsU["Width"].Formula = "=(1.0+2.5)"
        shape.CellsU["Height"].Formula = "=(0.0+1.5)"

        # BUILD UP THE REQUEST
        items = [
                Shape_GetFormulas_Record(IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowXFormOut, IVisio.VisCellIndices.visXFormWidth),
                Shape_GetFormulas_Record(IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowXFormOut, IVisio.VisCellIndices.visXFormHeight)
        ]

        # MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        SRCStream = get_new_system_array(System.Int16, len(items)*3)
        for i in xrange(len(items)) :
                SRCStream[i * 3 + 0] = items[i].SectionIndex
                SRCStream[i * 3 + 1] = items[i].RowIndex
                SRCStream[i * 3 + 2] = items[i].CellIndex

        # EXECUTE THE REQUEST
        formulas_sa = get_outref_to_system_array(System.Object) 
        SRCStream_sa = get_ref_to_system_array(System.Int16,SRCStream)  
        shape.GetFormulasU(SRCStream_sa, formulas_sa)

        # OUTPUT BACK TO SOMETHING USEFUL 
        formulas = get_new_system_array(System.String,formulas_sa.Length)
        formulas_sa.CopyTo(formulas, 0);

        shape.Text = System.String.Format("Formulas={0},{1}", formulas[0], formulas[1])

def Shape_GetResults( doc ):

        pages = doc.Pages
        page = pages.Add()
        page.NameU = "SGR"

        shape = page.DrawRectangle(1, 1, 4, 3)
        shape.CellsU["Width"].Formula = "=(1.0+2.5)"
        shape.CellsU["Height"].Formula = "=(0.0+1.5)"

        # BUILD UP THE REQUEST
        flags = System.Int16(IVisio.VisGetSetArgs.visGetFloats)
        items = [
                Shape_GetResults_Record(IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowXFormOut, IVisio.VisCellIndices.visXFormWidth, IVisio.VisUnitCodes.visNoCast),
                Shape_GetResults_Record(IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowXFormOut, IVisio.VisCellIndices.visXFormHeight, IVisio.VisUnitCodes.visNoCast)
        ]

        # MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        SRCStream = get_new_system_array(System.Int16, len(items)*3)
        unitcodes = get_new_system_array(System.Object, len(items))
        for i in xrange(len(items)) :
                SRCStream[i * 3 + 0] = items[i].SectionIndex
                SRCStream[i * 3 + 1] = items[i].RowIndex
                SRCStream[i * 3 + 2] = items[i].CellIndex
                unitcodes[i] = items[i].UnitCode

        # EXECUTE THE REQUEST
        results_sa = get_outref_to_system_array(System.Object) 
        SRCStream_sa = get_ref_to_system_array(System.Int16,SRCStream) 
        unitcodes_sa = get_ref_to_system_array(System.Object,unitcodes) 
        shape.GetResults(SRCStream_sa, flags, unitcodes_sa, results_sa)

        # OUTPUT BACK TO SOMETHING USEFUL 
        results = get_new_system_array(System.Double,results_sa.Length)
        results_sa.CopyTo(results, 0);

        shape.Text = System.String.Format("Results={0},{1}", results[0], results[1])

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
        SRCStream = get_new_system_array(System.Int16, len(items)*3)
        formulas = get_new_system_array(System.Object, len(items))
        for i in xrange(len(items)) :
                SRCStream[i * 3 + 0] = items[i].SectionIndex
                SRCStream[i * 3 + 1] = items[i].RowIndex
                SRCStream[i * 3 + 2] = items[i].CellIndex
                formulas[i] = items[i].Formula

        # EXECUTE THE REQUEST
        formulas_sa = get_ref_to_system_array(System.Object,formulas) 
        SRCStream_sa = get_ref_to_system_array(System.Int16,SRCStream)  
        count = shape.SetFormulas(SRCStream_sa, formulas_sa, flags)

        shape.Text = System.String.Format("SetFormulas")

def Shape_SetResults( doc ):

        pages = doc.Pages
        page = pages.Add()
        page.NameU = "SSR"
        
        shape = page.DrawRectangle(1, 1, 4, 3)
        shape.CellsU["Width"].Formula = "=(1.0+2.5)"
        shape.CellsU["Height"].Formula = "=(0.0+1.5)"

        # BUILD UP THE REQUEST
        flags = System.Int16(0)        
        items = [

                Shape_SetResults_Record( IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowXFormOut, IVisio.VisCellIndices.visXFormWidth, 8.0, IVisio.VisUnitCodes.visNoCast),                
                Shape_SetResults_Record( IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowXFormOut, IVisio.VisCellIndices.visXFormHeight, 1.0, IVisio.VisUnitCodes.visNoCast)
        ]

        # MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        SRCStream = get_new_system_array(System.Int16, len(items)*3)
        results = get_new_system_array(System.Object, len(items))
        unitcodes = get_new_system_array(System.Object, len(items))
        for i in xrange(len(items)) :
                SRCStream[i * 3 + 0] = items[i].SectionIndex
                SRCStream[i * 3 + 1] = items[i].RowIndex
                SRCStream[i * 3 + 2] = items[i].CellIndex
                results[i] = items[i].Result
                unitcodes[i] = items[i].UnitCode

        # EXECUTE THE REQUEST
        results_sa = get_ref_to_system_array(System.Object,results)
        unitcodes_sa = get_ref_to_system_array(System.Object,unitcodes)
        SRCStream_sa = get_ref_to_system_array(System.Int16,SRCStream)  
        count = shape.SetResults(SRCStream_sa, unitcodes_sa, results_sa, flags)

        shape.Text = System.String.Format("SetResults")

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
        SRCStream = get_new_system_array(System.Int16, len(items)*4)
        for i in xrange(len(items)) :
                SRCStream[i * 4 + 0] = items[i].ShapeID
                SRCStream[i * 4 + 1] = items[i].SectionIndex
                SRCStream[i * 4 + 2] = items[i].RowIndex
                SRCStream[i * 4 + 3] = items[i].CellIndex

        # EXECUTE THE REQUEST
        formulas_sa = get_outref_to_system_array(System.Object) 
        SRCStream_sa = get_ref_to_system_array(System.Int16,SRCStream)  
        page.GetFormulasU(SRCStream_sa, formulas_sa)

        # OUTPUT BACK TO SOMETHING USEFUL 
        formulas = get_new_system_array(System.String,formulas_sa.Length)
        formulas_sa.CopyTo(formulas, 0);

        shape.Text = System.String.Format("Formulas={0},{1}", formulas[0], formulas[1])


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
        SRCStream = get_new_system_array(System.Int16, len(items)*4)
        unitcodes = get_new_system_array(System.Object, len(items))
        for i in xrange(len(items)) :
                SRCStream[i * 4 + 0] = items[i].ShapeID
                SRCStream[i * 4 + 1] = items[i].SectionIndex
                SRCStream[i * 4 + 2] = items[i].RowIndex
                SRCStream[i * 4 + 3] = items[i].CellIndex
                unitcodes[i] = items[i].UnitCode

        # EXECUTE THE REQUEST
        results_sa = get_outref_to_system_array(System.Object) 
        SRCStream_sa = get_ref_to_system_array(System.Int16,SRCStream) 
        unitcodes_sa = get_ref_to_system_array(System.Object,unitcodes) 
        page.GetResults(SRCStream_sa, flags, unitcodes_sa, results_sa)

        # OUTPUT BACK TO SOMETHING USEFUL 
        results = get_new_system_array(System.Double,results_sa.Length)
        results_sa.CopyTo(results, 0);

        shape.Text = System.String.Format("Results={0},{1}", results[0], results[1])


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
        SRCStream = get_new_system_array(System.Int16, len(items)*4)
        formulas = get_new_system_array(System.Object, len(items))
        for i in xrange(len(items)) :
                SRCStream[i * 4 + 0] = items[i].ShapeID
                SRCStream[i * 4 + 1] = items[i].SectionIndex
                SRCStream[i * 4 + 2] = items[i].RowIndex
                SRCStream[i * 4 + 3] = items[i].CellIndex
                formulas[i] = items[i].Formula

        # EXECUTE THE REQUEST
        formulas_sa = get_ref_to_system_array(System.Object,formulas) 
        SRCStream_sa = get_ref_to_system_array(System.Int16,SRCStream)  
        count = page.SetFormulas(SRCStream_sa, formulas_sa, flags)

        shape.Text = System.String.Format("SetFormulas")

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
        SRCStream = get_new_system_array(System.Int16, len(items)*4)
        results = get_new_system_array(System.Object, len(items))
        unitcodes = get_new_system_array(System.Object, len(items))
        for i in xrange(len(items)) :
                SRCStream[i * 4 + 0] = items[i].ShapeID
                SRCStream[i * 4 + 1] = items[i].SectionIndex
                SRCStream[i * 4 + 2] = items[i].RowIndex
                SRCStream[i * 4 + 3] = items[i].CellIndex
                results[i] = items[i].Result
                unitcodes[i] = items[i].UnitCode

        # EXECUTE THE REQUEST
        results_sa = get_ref_to_system_array(System.Object,results)
        unitcodes_sa = get_ref_to_system_array(System.Object,unitcodes)
        SRCStream_sa = get_ref_to_system_array(System.Int16,SRCStream)  
        count = page.SetResults(SRCStream_sa, unitcodes_sa, results_sa, flags)

        shape.Text = System.String.Format("SetResults")

        
def test() :
    visapp = IVisio.ApplicationClass() 
    doc = visapp.Documents.Add("") 

    Shape_GetFormulas(doc)
    Shape_GetResults(doc)          
    Shape_SetFormulas(doc)
    Shape_SetResults(doc)
    Page_GetFormulas(doc)
    Page_GetResults(doc)
    Page_SetFormulas(doc)
    Page_SetResults(doc)

def get_new_system_array(T,length) :
    array = System.Array.CreateInstance(T, length)
    return array

def get_ref_to_system_array(T,array) :
    ref = clr.Reference[System.Array[T]](array) 
    return ref

def get_outref_to_system_array(T) :
    ref = clr.Reference[System.Array[T]](System.Array[T]([])) 
    return ref
        
if __name__ == "__main__" :
        test()
        
