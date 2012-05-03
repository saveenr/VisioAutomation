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

class DOMConnectionType:
    Manual = 0
    Auto = 1

class DOMConnection(object):

    def __init__(self , fromshape, toshape, connectorshape,contype) :
        self.FromShape = fromshape
        self.ToShape = toshape
        self.ConnectorShape = connectorshape
        self.Type = contype
        self.Direction = 0

class DOM(object): 
    
    def __init__( self ) :
        self.Shapes = []
        self.Stencils = []
        self.Masters = []
        self.Connections = []

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
        con = DOMConnection(fromshape, toshape, connectorshape, DOMConnectionType.Manual)
        self.Connections.append(con)

    def AutoConnect( self, fromshape, toshape, connectorshape, direction=0) :
        con = DOMConnection(fromshape, toshape, connectorshape, DOMConnectionType.Auto)
        con.Direction = direction
        self.Connections.append(con)

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
        # Visio 2010 Shape.AutoConnect on MSDN http://msdn.microsoft.com/en-us/library/ff765915.aspx
        # Visio 2010 Connectivity APIs: http://blogs.msdn.com/b/visio/archive/2009/09/22/the-visio-2010-connectivity-api.aspx
        # Visio 2010 Page.AutoConnectMany http://msdn.microsoft.com/en-us/library/ff765694.aspx


        
        nonbatch_connects = []
        batch_autoconnects = []
        for i,cxn in enumerate( self.Connections ) :
            if (cxn.Type == DOMConnectionType.Manual) :
                nonbatch_connects.append(cxn)
            elif (cxn.Type == DOMConnectionType.Auto and cxn.FromShape.VisioShape == cxn.ToShape.VisioShape) :
                nonbatch_connects.append(cxn)
            elif (cxn.Type == DOMConnectionType.Auto and cxn.FromShape.VisioShape != cxn.ToShape.VisioShape) :
                batch_autoconnects.append(cxn)

        if (len(nonbatch_connects)>0):
            for i,cxn in enumerate( nonbatch_connects ) :
                self.__connect(cxn.FromShape.VisioShape, cxn.ToShape.VisioShape, cxn.ConnectorShape.VisioShape)

        if (len(batch_autoconnects)>0):
            fromshapeids =[]
            toshapeids=[]
            placementdirs = []
            connectors = []
            for cxn in batch_autoconnects:
                fromshapeids.append( cxn.FromShape.VisioShapeID )
                toshapeids.append( cxn.ToShape.VisioShapeID )
                placementdirs.append( cxn.Direction )
                connectors.append(None)
                #autoconnectshape = cxn.FromShape.VisioShape.AutoConnect( cxn.ToShape.VisioShape, cxn.Direction, cxn.ConnectorShape.VisioShape )                
            page.AutoConnectMany(fromshapeids, toshapeids, placementdirs, None )

    def __connect( self, fromshape, toshape, connectorshape ) :
        cxn_from_beginx = connectorshape.CellsU( "BeginX" )
        cxn_to_endy = connectorshape.CellsU( "EndY" )
        cxn_from_beginx.GlueTo(fromshape.CellsSRC(1, 1, 0)) 
        cxn_to_endy.GlueTo(toshape.CellsSRC(1, 1, 0))
