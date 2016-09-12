# -*- coding: utf-8 -*-

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

# salecting and setting text

visio.client.Selection.SelectAll() 
targets = VisioAutomation.Scripting.TargetShapes()
visio.client.Text.Set( targets,  [ "A", "B", "C" ] )

# drawing a grid the manual way (slow)

visio.client.Page.New(pagesize, False)
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


#cells= [VisioAutomation.ShapeSheet.SRCConstants.FillPattern for i in patterns]
#formulas = [ str(i) for i in patterns ]
#getsetargs = 0
#print [ int(s.ID16) for s in shapes ]
#targets2 = VisioAutomation.Scripting.TargetShapes(shapes)
#visio.client.ShapeSheet.SetFormula(targets2, cells , formulas, getsetargs)  


# drawing tabular data
# drawing hierrchical data
# using mSAGL



"""



>>> vi.Page.New() 
>>> patterns = range(41) 
>>> for (i,pattern) in enumerate(patterns): 
>>>     vi.Draw.Rectangle( i , 0, i+1, 0+1 ) 
>>>     vi.Fill.Pattern = pattern 
>>> vi.Page.ResizeToFitContents()

Demo 2 – Drawing Tabular Data
The interactive shell extensively uses System.Data.DataTable to store tabular data
>>> data = ( ('Hello',1) , ('World',2) ) 
>>> datatable = ToDataTable( data ) 
>>> vi.Draw.Table( datatable )

We can get data from a CSV file
create a CSV file in Excel

And then load it as a DataTable and let Visio Draw it
>>> datatable = vi.Data.ImportCSV( r"D:\saveenr\data1.csv" ) 
>>> vi.Draw.Table( datatable )

Of course, you can load an XLSX file. In this case, you’ll have to identify the name of the worksheet also…
>>> datatable = vi.Data.ImportExcelWorksheet( r"d:\\data1.xlsx" , "Sheet1" ) 
>>> vi.Draw.Table( datatable )
Demo 3 – Drawing Hierarchical Data
Let’s draw a tree from a set of directory paths
vi.Page.New() 
items=List[System.String]() 
items.Add( r'c:' ) 
items.Add( r'c:\windows' ) 
items.Add( r'c:\windows\system' ) 
items.Add( r'c:\windows\system32' ) 
items.Add( r'c:\windows\tasks' ) 
items.Add( r'c:\program files' ) 
items.Add( r'c:\program files\office live' ) 
items.Add( r'c:\baz' ) 
dir = Isotope.Drawing.CardinalDirection.Down 
tree = vi.Data.PathsToTree( items ) 
vi.Draw.Tree( tree.Root , dir  ) 
vi.Page.ResizeToFitContents() 
vi.Zoom.ToPage()

Demo 4 – Automatic Layout Directed Graphs
Technical Note: the AutoLayoutDrawing class uses Microsoft’s Automatic Graph Layout library
d= VisioDOM.AutoLayout.AutoLayoutDrawing() 
s1= d.AddShape('A') 
s2= d.AddShape('B') 
s3= d.AddShape('C') 
c1 = d.Connect('c1',s1,s2) 
c2 = d.Connect('c2',s1,s3) 
r = VisioDOM.AutoLayout.AutoLayoutRenderer() 
r.RenderToVisio( d , vi.visapp , True) 
vi.Zoom.ToPage()

This was a simple example, you can draw much more complicated diagrams using the autolayout feature
"""


