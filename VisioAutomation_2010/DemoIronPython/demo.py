# -*- coding: utf-8 -*-

import visio


visio.client.Application.New()
visio.client.Document.New(8.5,11)

visio.client.Draw.Rectangle(0,0,1,1)
visio.client.Draw.Oval(1,1,2,2)
visio.client.Draw.Line(2,2,3,3  )

    
# basic drawing
# droppping a master
# sleecting and setting text
# drawing all the fill patterns on a page
# drawing tabular data
# drawing hierrchical data
# using mSAGL



"""

Draw some simple shapes
>>> vi.Draw.Rectangle( 0, 0, 1,1 ) 
>>> vi.Draw.Oval( 2, 2, 3,3 ) 
>>> vi.Draw.Line( 4, 4, 5,5 )

Drop a master
>>> vi.Drop.Master( "BASIC_U.VSS", "Rectangle", 2, 5 )

Set the text of all the shapes
>>> vi.Select.All() 
>>> vi.Text.PlainText = "Hello World""

Lets draw all the fill patterns on a new page
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


