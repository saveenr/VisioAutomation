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
       
