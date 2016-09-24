import visio
import VisioAutomation

import System
# Create a new application and document


visio.client.Application.New()
pagesize = VisioAutomation.Drawing.Size(8.5,11)
visio.client.Document.New(pagesize)


basic_stencil = visio.client.Document.OpenStencil("Basic_U.VSS")
master = visio.client.Master.Get("Rectangle", basic_stencil)
visio.client.Master.Drop(master, VisioAutomation.Drawing.Point(0,0) )
visio.client.Master.Drop(master, VisioAutomation.Drawing.Point(2,2) )
visio.client.Master.Drop(master, VisioAutomation.Drawing.Point(6,6) )

# dropping multiple masters (fast)

visio.client.Page.New(pagesize, False)
points=[]
points.append(VisioAutomation.Drawing.Point(0,0))
points.append(VisioAutomation.Drawing.Point(2,2))
points.append(VisioAutomation.Drawing.Point(6,6))
masters = [ master ]
visio.client.Master.Drop(masters,points)

# selecting and setting text

visio.client.Selection.SelectAll() 
targets = VisioAutomation.Scripting.TargetShapes()
visio.client.Text.Set( targets,  [ "Foo", "Bar", "Beer" ] )

dic = System.Collections.Hashtable()
dic [ "PinX" ] = 1.0
dic [ "FillForegnd" ] = "rgb(255,255,0)"
visio.client.ShapeSheet.SetShapeCells( targets, dic, False, False)
