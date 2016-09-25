import visio
import VisioAutomation

# Create a new application and document

visio.client.Application.New()
pagesize = VisioAutomation.Drawing.Size(8.5,11)
visio.client.Document.New(pagesize)

# Documents always have at least one page
# so we can begin drawing now. We'll start
# with some basic shapes

visio.client.Draw.Rectangle(0,0,1,1)
visio.client.Draw.Oval(1,1,2,2)
visio.client.Draw.Line(2,2,3,3  )

# Let's start with a new page


visio.client.Page.New(pagesize, False)

# Instead of drawing shapes, in Visio we
# "drop" shapes from "master shapes" that
# are in "stencils". We'll load a stencil
# and then drop a master from that stencil

basic_stencil = visio.client.Document.OpenStencil("Basic_U.VSS")
master = visio.client.Master.Get("Rectangle", basic_stencil)
visio.client.Master.Drop(master, VisioAutomation.Drawing.Point(0,0) )
visio.client.Master.Drop(master, VisioAutomation.Drawing.Point(2,2) )
visio.client.Master.Drop(master, VisioAutomation.Drawing.Point(6,6) )

# dropping multiple masters (fast)

visio.client.Page.New(pagesize, False)
points=[]
points.append(VisioAutomation.Drawing.Point(0,0))
points.append( VisioAutomation.Drawing.Point(2,2))
points.append(VisioAutomation.Drawing.Point(6,6))
masters = [ master ]
visio.client.Master.Drop(masters,points)

# selecting and setting text

visio.client.Selection.SelectAll() 
targets = VisioAutomation.Scripting.TargetShapes()
visio.client.Text.Set( targets,  [ "A", "B", "C" ] )

