from __future__ import division
import array
import sys 

import win32com.client 
win32com.client.gencache.EnsureDispatch("Visio.Application") 

import ShapeSheet
import Errors
import Drawing

class DOMShape(object):
    
    def __init__( self , master, pos) :
        self.Master = master
        self.Cells = { } 
        self.VisioShape = None
        self.VisioShapeID = None
        self.Text = None
        self.DropPos = None

        if ( isinstance(pos,Drawing.Point) ) :
            self.DropPosition = pos
            self.DropSize = None
        elif ( isinstance(pos,Drawing.Rectangle) ) :
            self.DropSize = pos.Size
            self.DropPosition = pos.CenterPoint
        else :
            raise Errors.VisioPyError()

class DOMConnection(DOMShape):

    def __init__(self , fromshape, toshape, master) :
        super(DOMConnection, self).__init__(master, Drawing.Point(-1,-1)) 
        self.FromShape = fromshape
        self.ToShape = toshape

class DOMMaster(object):

    def __init__(self , mastername, stencil) :
        self.MasterName = mastername
        self.StencilName = stencil
        self.VisioMaster = None

class DOM(object): 
    
    def __init__( self ) :
        self.Shapes = []
        self.Stencils = []
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

    def Connect( self, fromshape, toshape, connecterobject, text=None , cells = None) :
        if (not isinstance(connecterobject,DOMMaster)) :    
            raise Errors.VisioPyError()

        cxn = DOMConnection(fromshape, toshape, connecterobject)
        if (cells!=None) : cxn.Cells = cells
        cxn.Text = text
        self.Connections.append(cxn)
        return cxn

    def OpenStencil( self, name) :
        stencil = DOMStencil(name)
        self.Stencils.append( stencil )
        return stencil

    def Render( self, page ) :

        with PerfContext(page.Application) :
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
            masters = set( s.Master for s in self.Shapes if s.Master != None)
            for cxn in self.Connections :
                masters.add( cxn.Master )

            master_cache = {}
            for master in masters:
                stencildoc = stencil_cache[ master.StencilName.lower() ]
                mastername = master.MasterName.lower()
                vmaster = master_cache.get( mastername, None )
                if (vmaster == None) :
                    vmaster = stencildoc.Masters.ItemU(master.MasterName) 
                master.VisioMaster = vmaster 
                if (master.VisioMaster==None) :
                    raise Errors.VisioPyError()

            # Perform the basic drop of all the shapes
            vmasters = []
            xyarray = []
            for shape in self.Shapes:
                vmasters.append( shape.Master.VisioMaster )
                xyarray.append( shape.DropPosition.X )
                xyarray.append( shape.DropPosition.Y )
            num_shapes,shape_ids = page.DropMany( vmasters, xyarray) 


            # Ensure that we have stored the corresponding shape object and shapeid for each dropped object
            page_shapes = page.Shapes
            for i,shape in enumerate( self.Shapes ) :
                shape.VisioShapeID = shape_ids[i]
                shape.VisioShape = page_shapes.ItemFromID( shape_ids[i] )

            #set any shape properties
            u = ShapeSheet.Update()
            for shape in self.Shapes:
                if (shape.DropSize!=None):
                    u.Add( shape.VisioShapeID, ShapeSheet.SRCConstants.Width , shape.DropSize.Width)
                    u.Add( shape.VisioShapeID, ShapeSheet.SRCConstants.Width , shape.DropSize.Height)
                if (len(shape.Cells)>0) :
                    for src in shape.Cells :
                        formula = shape.Cells[src]
                        u.Add( shape.VisioShapeID, src, formula)
                    
            result = u.SetFormulas(page) 
        

            self.__connectshapes(page)

            #set any connection properties
            u = ShapeSheet.Update()
            for cxn in self.Connections:
                shape = cxn.VisioShape
                shapeid = cxn.VisioShapeID
                if (len(cxn.Cells)>0) :
                    for src in cxn.Cells :
                        formula = cxn.Cells[src]
                        u.Add( cxn.VisioShapeID, src, formula)
            result = u.SetFormulas(page) 

            # Set the text for shapes and connections
            for shape in self.Shapes:
                if (shape.Text != None and shape.Text!='') :
                    shape.VisioShape.Text = shape.Text

            for cxn in self.Connections:
                if (cxn.Text != None and cxn.Text!='') :
                    cxn.VisioShape.Text = cxn.Text


    def __connectshapes( self , page ) :

        # Finally perform the connections
        if (len(self.Connections)<1) : 
            return

        # Drop all the masters
        vmasters = []
        xyarray = []
        for cxn in self.Connections:
            vmasters.append( cxn.Master.VisioMaster )
            xyarray.append( -2 )
            xyarray.append( -2 )
        num_shapes,shape_ids = page.DropMany( vmasters, xyarray) 
        vshapes = [page.Shapes.ItemFromID( id ) for id in shape_ids]

        # Connect them
        direction = 0
        for i,cxn in enumerate(self.Connections):
            fromshape = cxn.FromShape.VisioShape
            toshape = cxn.ToShape.VisioShape
            vmaster = cxn.Master.VisioMaster
            connectorshape = vshapes[i]

            cxn_from_beginx = connectorshape.CellsU( "BeginX" )
            cxn_to_endy = connectorshape.CellsU( "EndY" )
            cxn_from_beginx.GlueTo(fromshape.CellsSRC(1, 1, 0)) 
            cxn_to_endy.GlueTo(toshape.CellsSRC(1, 1, 0))

            cxn.VisioShape = connectorshape
            cxn.VisioShapeID = cxn.VisioShape.ID

class UndoContext:

    def __init__(self, app, name):
        self.Application = app
        self.ScopeID =None
        self.Name = name
        self.Commit = None

    def __enter__(self):
        self.ScopeID = self.Application.BeginUndoScope(self.Name);
        self.Commit = True

    def __exit__(self, type, value, tb):
        self.Application.EndUndoScope(self.ScopeID, self.Commit )
        del self.Application

class PerfSettings:

    def __init__(self):
        self.EnableAutoConnect=None
        self.LiveDynamics=None
        self.ScreenUpdating=None
        self.DeferRecalc=None

    def Load(self,app):
        self.old_livedynamics = app.LiveDynamics
        self.old_autoconnect = app.Settings.EnableAutoConnect
        self.old_deferrecalc = app.DeferRecalc
        self.old_screenupdating = app.ScreenUpdating

    def Apply(self,app):
        if (self.ScreenUpdating!=None) : app.ScreenUpdating = self.ScreenUpdating
        if (self.DeferRecalc!=None) : app.DeferRecalc = self.DeferRecalc 
        if (self.EnableAutoConnect!=None) : app.Settings.EnableAutoConnect = self.EnableAutoConnect 
        if (self.LiveDynamics!=None) : app.LiveDynamics = self.LiveDynamics 

    def __exit__(self, type, value, tb):
        self.app.ScreenUpdating = self.old_screenupdating;
        self.app.DeferRecalc = self.old_deferrecalc;
        self.app.Settings.EnableAutoConnect = self.old_autoconnect;
        self.app.LiveDynamics = self.old_livedynamics;

class PerfContext:

    def __init__(self, app):
        self.Application = app

        self.OldSettings = PerfSettings()

        ps = PerfSettings()
        ps.ScreenUpdating = 0
        ps.DeferRecalc = 1
        ps.EnableAutoConnect = False
        ps.LiveDynamics = False

        self.NewSettings = ps

    def __enter__(self):
        self.OldSettings.Load(self.Application)
        self.NewSettings.Apply(self.Application)

    def __exit__(self, type, value, tb):
        self.OldSettings.Apply(self.Application)
        del self.Application