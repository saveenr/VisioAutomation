from __future__ import division
import array
import sys 
import win32com.client 
win32com.client.gencache.EnsureDispatch("Visio.Application") 

from Drawing import *
from ShapeSheet import *

class DOMShape:
    
    def __init__( self , mastername, stencilname, pos) :
        self.MasterName = mastername
        self.StencilName = stencilname
        self.Master = None
        if ( isinstance(pos,Point) ) :
            self.DropPosition = pos
            self.DropSize = None
        elif ( isinstance(pos,Rectangle) ) :
            self.DropSize = ((pos.Right-pos.Left),(pos.Top-pos.Bottom))
            self.DropPosition = Point( (pos.Right-pos.Left)/2 , (pos.Top-pos.Bottom)/2 )
        else :
            print ">>>", pos is Rectangle
            raise DOM()
            #raise some error
        self.VisioShape = None
        self.VisioShapeID = None
        self.Text = None
        

def openstencilx(docs, stencilname) :
    stencildocflags = win32com.client.constants.visOpenRO | win32com.client.constants.visOpenDocked 
    stencildoc = docs.OpenEx(stencilname , stencildocflags )
    return stencildoc

class DOM : 
    
    def __init__( self ) :
        self.Shapes = []
        self.Connections = []

    def Drop( self, mastername, stencilname, pos , text=None) :
        domshape = DOMShape( mastername, stencilname, pos )
        domshape.Text = text
        self.Shapes.append(domshape) 
        return domshape

    def Connect( self, fromshape, toshape, connectorshape ) :
        self.Connections.append((fromshape, toshape, connectorshape))

    def Render( self, page ) :
        # Load all the stencils
        # Goal: prevent trying to reload the same stencil multiple times
        # Goal: minimize having to use COM to lookup stencil documents by name
        docs = page.Application.Documents
        stencilnames = set(s.StencilName.lower() for s in self.Shapes)
        stencil_cache = {}
        for stencilname in stencilnames:
            stencildoc = openstencilx( docs, stencilname )       
            stencil_cache[ stencilname ] = stencildoc 

        # cache all the master references
        # Goal: minimize having to use COM to lookup master objects by name
        master_cache = {}
        for shape in self.Shapes:
            stencildoc = stencil_cache[ shape.StencilName.lower() ]
            mastername = shape.MasterName.lower()
            master = master_cache.get( mastername, None )
            if (master == None) :
                master = stencildoc.Masters.ItemU(shape.MasterName) 
            shape.Master = master

        # Perform the basic drop of all the masters
        masters = []
        xyarray = []
        for shape in self.Shapes:
            masters.append( shape.Master )
            xyarray.append( shape.DropPosition.X )
            xyarray.append( shape.DropPosition.Y )
        num_shapes,shape_ids = page.DropMany( masters, xyarray) 


        # Ensure that we have stored the corresponding shape object and shapeid for each dropped object
        page_shapes = page.Shapes
        for i,shape in enumerate( self.Shapes ) :
            shape.VisioShapeID = shape_ids[i]
            shape.VisioShape = page_shapes.ItemFromID( shape_ids[i] )

        #set any dropsizes
        u = Update()
        for shape in self.Shapes:
            if (shape.DropSize!=None):
                u.Add( shape.VisioShapeID, SRCConstants.Width , str(shape.DropSize[0]))
                u.Add( shape.VisioShapeID, SRCConstants.Width , str(shape.DropSize[1]))
        result = u.SetFormulas(page) 
        
        for shape in self.Shapes:
            if (shape.Text != None and shape.Text!='') :
                shape.VisioShape.Text = shape.Text

        # Finally perform the connections
        for i,cxn in enumerate( self.Connections ) :
            self.__connect(cxn[0].VisioShape, cxn[1].VisioShape, cxn[2].VisioShape)


    def __connect( self, fromshape, toshape, connectorshape ) :
        cxn_from_beginx = connectorshape.CellsU( "BeginX" )
        cxn_to_endy = connectorshape.CellsU( "EndY" )
        cxn_from_beginx.GlueTo(fromshape.CellsSRC(1, 1, 0)) 
        cxn_to_endy.GlueTo(toshape.CellsSRC(1, 1, 0))
