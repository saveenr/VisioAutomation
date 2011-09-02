import clr 
import System 
import functools
clr.AddReference("Microsoft.Office.Interop.Visio") 
import Microsoft.Office.Interop.Visio 
IVisio = Microsoft.Office.Interop.Visio


def get_new_system_array(T,length) :
    array = System.Array.CreateInstance(T, length)
    return array

def get_ref_to_system_array(T,array) :
    ref = clr.Reference[System.Array[T]](array) 
    return ref

def get_outref_to_system_array(T) :
    ref = clr.Reference[System.Array[T]](System.Array[T]([])) 
    return ref
        
        


def dropmany( page, masters, xys ) :
        masters_obj_arr = System.Array[object]( masters ) 
        xys = System.Array[System.Double]( xys ) 
        out_ids = get_ref_short_array() 
        page.DropManyU( masters_obj_arr, xys, out_ids ) 
        return out_ids

def set_formulas( page, items ) :
        # BUILD UP THE REQUEST
        flags = System.Int16(IVisio.VisGetSetArgs.visSetBlastGuards | IVisio.VisGetSetArgs.visSetUniversalSyntax)        

        # MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        SRCStream = get_new_system_array(System.Int16, len(items)*4)
        formulas = get_new_system_array(System.Object, len(items))
        for i in xrange(len(items)) :
                item = items[i]
                shapeid = item[0]
                src = item[1]
                formula = item[2]

                SRCStream[i * 4 + 0] = System.Int16(shapeid         )
                SRCStream[i * 4 + 1] = System.Int16(src.Section     )
                SRCStream[i * 4 + 2] = System.Int16(src.Row         )
                SRCStream[i * 4 + 3] = System.Int16(src.Cell        )
                formulas[i] = formula

        # EXECUTE THE REQUEST
        formulas_sa = get_ref_to_system_array(System.Object,formulas) 
        SRCStream_sa = get_ref_to_system_array(System.Int16,SRCStream)  
        count = page.SetFormulas(SRCStream_sa, formulas_sa, flags)

class SRC :
    def __init__(self, s,r,c ) :
        self.Section = s
        self.Row = r
        self.Cell = c

    @staticmethod
    def GetSRCBuilder_SR( s, r , c ) :
        return functools.partial( SRC, *{s:s,r:r} )

class SIDSRC :
    def __init__(self, sid, src) :
        self.ShapeID = sid
        self.SRC = src

class SRCConstants :

    fillsrc_builder = SRC.GetSRCBuilder_SR( IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowFill, IVisio.VisCellIndices.visFillForegnd )
    linesrc_builder = SRC.GetSRCBuilder_SR( IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowLine, IVisio.VisCellIndices.visFillForegnd )
    charsrc_builder = SRC.GetSRCBuilder_SR( IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowCharacter, IVisio.VisCellIndices.visFillForegnd )
    charsrc_builder = SRC.GetSRCBuilder_SR( IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowParagraph, IVisio.VisCellIndices.visFillForegnd )
    
    FillBkgnd = fillsrc_builder( IVisio.VisCellIndices.visFillBkgnd )
    FillBkgndTrans = fillsrc_builder( IVisio.VisCellIndices.visFillBkgndTrans)
    FillForegnd = fillsrc_builder( IVisio.VisCellIndices.visFillForegnd )
    FillForegndTrans = fillsrc_builder( IVisio.VisCellIndices.visFillForegndTrans )
    FillPattern = fillsrc_builder( IVisio.VisCellIndices.visFillPattern)


    LineColor = fillsrc_builder( IVisio.VisCellIndices.visLineColor)
    LineColorTrans = linesrc_builder( IVisio.VisCellIndices.visLineColorTrans)
    LinePattern = linesrc_builder( IVisio.VisCellIndices.visLinePattern)
    LineWeight = linesrc_builder( IVisio.VisCellIndices.visLineWeight)
