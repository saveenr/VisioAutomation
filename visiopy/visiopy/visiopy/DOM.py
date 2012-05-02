import array
import sys 
import win32com.client 
win32com.client.gencache.EnsureDispatch("Visio.Application") 

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
            self.__connect(cxn[0].VisioShape, cxn[1].VisioShape, cxn[2].VisioShape)


    def __connect( self, fromshape, toshape, connectorshape ) :
        cxn_from_beginx = connectorshape.CellsU( "BeginX" )
        cxn_to_endy = connectorshape.CellsU( "EndY" )
        cxn_from_beginx.GlueTo(fromshape.CellsSRC(1, 1, 0)) 
        cxn_to_endy.GlueTo(toshape.CellsSRC(1, 1, 0))
