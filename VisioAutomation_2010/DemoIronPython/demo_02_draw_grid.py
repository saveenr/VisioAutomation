import visio
import VisioAutomation

# Create a new application and document

visio.client.Application.New()
pagesize = VisioAutomation.Drawing.Size(8.5,11)
visio.client.Document.New(pagesize)


shapes = []
n = 5
for i in xrange(n*n): 
    left = i % n
    bottom = i / n
    shape = visio.client.Draw.Rectangle( left, bottom, left+1, bottom+1 ) 
    shapes.append(shape)
bordersize = VisioAutomation.Drawing.Size(1,1)
zoom_to_page = True
visio.client.Page.ResizeToFitContents(bordersize,zoom_to_page)


# drawing a grid the fast way with a grid layout - makes the overall task simpler
visio.client.Page.New(pagesize, False)
basic_stencil = visio.client.Document.OpenStencil("Basic_U.VSS")
master = visio.client.Master.Get("Rectangle", basic_stencil)
cellsize = VisioAutomation.Drawing.Size(1,1)
gridlayout = VisioAutomation.Models.Layouts.Grid.GridLayout(5, 5, cellsize, master)
visio.client.Page.ResizeToFitContents(bordersize,zoom_to_page)


