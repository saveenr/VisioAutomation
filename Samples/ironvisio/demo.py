from ironvisio import *
import charting

app = IVisio.ApplicationClass()
docs = app.Documents
doc = docs.Add("")
page = app.ActivePage

values = [5,2,3,7,4]
category_labels = ["A", "B", "C", "D", "E"]

chart1= charting.VerticalBarChart()
chart1.DataPoints = [ charting.DataPoint(v) for v in values ]
chart1.Categories = category_labels 
chart1.Origin = charting.Point(0.5,0)

chart2= charting.CircleChart()
chart2.DataPoints = [ charting.DataPoint(v) for v in values ]
chart2.Categories = category_labels 
chart2.Origin = charting.Point(0.5,4)


chart1.Draw(page)
chart2.Draw(page)