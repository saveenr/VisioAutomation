from __future__ import division
import array
import sys 
import win32com.client 
win32com.client.gencache.EnsureDispatch("Visio.Application") 

from Drawing import *
from ShapeSheet import *

class DOMShape(object):
    
    def __init__( self , master, pos) :
        self.Master = master
        self.VisioMaster = None
        self.Cells = { } 

        if ( isinstance(pos,Point) ) :
            self.DropPosition = pos
            self.DropSize = None
        elif ( isinstance(pos,Rectangle) ) :
            self.DropSize = pos.Size
            self.DropPosition = pos.CenterPoint
        else :
            print ">>>", pos is Rectangle
            raise DOM()
            #raise some error
        self.VisioShape = None
        self.VisioShapeID = None
        self.Text = None
       

class DOMMaster(object):

    def __init__(self , mastername, stencil) :
        self.MasterName = mastername
        self.StencilName = stencil

class DOM(object): 
    
    def __init__( self ) :
        self.Shapes = []
        self.Connections = []
        self.Stencils = []
        self.Masters = []
        self.AutoConnections = []

    def Master( self, mastername, stencilname ) :
        m = DOMMaster( mastername, stencilname )
        return m

    def Drop( self, master, pos , text=None, cells=None) :
        domshape = DOMShape( master, pos )
        domshape.Text = text
        if (cells!=None) :
            domshape.Cells = cells
        self.Shapes.append(domshape) 
        return domshape

    def Connect( self, fromshape, toshape, connectorshape ) :
        self.Connections.append((fromshape, toshape, connectorshape))

    def AutoConnect( self, fromshape, toshape, connectorshape, direction) :
        self.Connections.append((fromshape, toshape, connectorshape, direction))

    def OpenStencil( self, name) :
        stencil = DOMStencil(name)
        self.Stencils.append( stencil )
        return stencil

    def Render( self, page ) :
        # Load all the stencils
        # Goal: prevent trying to reload the same stencil multiple times
        # Goal: minimize having to use COM to lookup stencil documents by name
        docs = page.Application.Documents
        stencilnames = set(s.Master.StencilName.lower() for s in self.Shapes)
        stencil_cache = {}
        stencildocflags = win32com.client.constants.visOpenRO | win32com.client.constants.visOpenDocked 
        for stencilname in stencilnames:
            stencildoc = docs.OpenEx(stencilname , stencildocflags )
            stencil_cache[ stencilname ] = stencildoc 

        # cache all the master references
        # Goal: minimize having to use COM to lookup master objects by name
        master_cache = {}
        for shape in self.Shapes:
            stencildoc = stencil_cache[ shape.Master.StencilName.lower() ]
            mastername = shape.Master.MasterName.lower()
            vmaster = master_cache.get( mastername, None )
            if (vmaster == None) :
                vmaster = stencildoc.Masters.ItemU(shape.Master.MasterName) 
            shape.VisioMaster = vmaster

        # Perform the basic drop of all the masters
        vmasters = []
        xyarray = []
        for shape in self.Shapes:
            vmasters.append( shape.VisioMaster )
            xyarray.append( shape.DropPosition.X )
            xyarray.append( shape.DropPosition.Y )
        num_shapes,shape_ids = page.DropMany( vmasters, xyarray) 


        # Ensure that we have stored the corresponding shape object and shapeid for each dropped object
        page_shapes = page.Shapes
        for i,shape in enumerate( self.Shapes ) :
            shape.VisioShapeID = shape_ids[i]
            shape.VisioShape = page_shapes.ItemFromID( shape_ids[i] )

        #set any dropsizes
        u = Update()
        for shape in self.Shapes:
            if (shape.DropSize!=None):
                u.Add( shape.VisioShapeID, SRCConstants.Width , shape.DropSize.Width)
                u.Add( shape.VisioShapeID, SRCConstants.Width , shape.DropSize.Height)
            if (len(shape.Cells)>0) :
                for src in shape.Cells :
                    formula = shape.Cells[src]
                    u.Add( shape.VisioShapeID, src, formula)
                    
        result = u.SetFormulas(page) 
        
        for shape in self.Shapes:
            if (shape.Text != None and shape.Text!='') :
                shape.VisioShape.Text = shape.Text

        # Finally perform the connections
        for i,cxn in enumerate( self.Connections ) :
            self.__connect(cxn[0].VisioShape, cxn[1].VisioShape, cxn[2].VisioShape)

        for i,cxn in enumerate( self.AutoConnections ) :
            # Shape.AutoConnect on MSDN http://msdn.microsoft.com/en-us/library/ff765915.aspx
            from_shape = cxn[0].VisioShape
            to_shape = cxn[1].VisioShape
            connectorshape = cxn[2]
            direction = cxn[3]
            autoconnectshape = from_shape.AutoConnect( to_shape, direction, connectorshape )


    def __connect( self, fromshape, toshape, connectorshape ) :
        cxn_from_beginx = connectorshape.CellsU( "BeginX" )
        cxn_to_endy = connectorshape.CellsU( "EndY" )
        cxn_from_beginx.GlueTo(fromshape.CellsSRC(1, 1, 0)) 
        cxn_to_endy.GlueTo(toshape.CellsSRC(1, 1, 0))
