import array
import sys 
import win32com.client 
win32com.client.gencache.EnsureDispatch("Visio.Application") 

from ShapeSheet import *
ShapeSheet.win32com = win32com

def openstencil(docs, stencilname) :
    stencildocflags = win32com.client.constants.visOpenRO | win32com.client.constants.visOpenDocked 
    stencildoc = docs.OpenEx(stencilname , stencildocflags )
    return stencildoc

def connect( fromshape, toshape, connectorshape ) :
    cxn_from_beginx = connectorshape.CellsU( "BeginX" )
    cxn_to_endy = connectorshape.CellsU( "EndY" )
    cxn_from_beginx.GlueTo(fromshape.CellsSRC(1, 1, 0)) 
    cxn_to_endy.GlueTo(toshape.CellsSRC(1, 1, 0))

def build_sidsrcstream( id_srcs ) :
    stream = []
    for id,src in id_srcs:
        stream.append(id)
        stream.append(src.Section)
        stream.append(src.Row)
        stream.append(src.Cell)
    return stream

class Query :

    def __init__(self) :
        self.items = []

    def Add(self, id, src) :
        item = (id,src)
        self.items.append(item)

    def GetFormulas(self, page) :
        stream = build_sidsrcstream( ( (id,src) for (id,src) in self.items )  )
        formulas = page.GetFormulas(stream)
        return formulas

    def GetResults(self, page) :
        stream = build_sidsrcstream( ( (id,src) for (id,src) in self.items )  )
        result = page.GetResults(stream,0,None)
        return result

class Update :

    def __init__(self) :
        self.items = []
        self.Flags = 0

    def Add(self, id, src, formula ) :
        item = (id,src,formula)
        self.items.append(item)

    def SetFormulas(self, page) :
        stream = build_sidsrcstream( ( (id,src) for (id,src,formula) in self.items )  )
        formulas = []
        for (id,src,formula) in self.items :
            formulas.append(formula)
        result = page.SetFormulas(stream, formulas, self.Flags)
        return result


class Point:
    
    def __init__( self , x, y) :
        self.X = x
        self.Y = y

class DOMShape:
    
    def __init__( self , master, pos) :
        self.Master = master
        self.DropPosition = pos
        self.VisioShape = None
        self.VisioShapeID = None

class DOM : 

    
    def __init__( self ) :
        self.Shapes = []
        self.Connections = []

    def Drop( self, master, pos ) :
        domshape = DOMShape( master, pos )
        self.Shapes.append(domshape) 
        return domshape

    def Connect( self, fromshape, toshape, connectorshape ) :
        self.Connections.append((fromshape, toshape, connectorshape))

    def Render( self, page ) :
        masters = []
        xyarray = []
        for shape in self.Shapes:
            masters.append( shape. Master )
            xyarray.append( shape.DropPosition.X )
            xyarray.append( shape.DropPosition.Y )
        num_shapes,shape_ids = page.DropMany( masters, xyarray) 
 
        page_shapes = page.Shapes
        for i,shape in enumerate( self.Shapes ) :
            shape.VisioShapeID = shape_ids[i]
            shape.VisioShape = page_shapes.ItemFromID( shape_ids[i] )

        for i,cxn in enumerate( self.Connections ) :
            connect(cxn[0].VisioShape, cxn[1].VisioShape, cxn[2].VisioShape)

if (__name__=='__main__') :
    pass
else :
    pass