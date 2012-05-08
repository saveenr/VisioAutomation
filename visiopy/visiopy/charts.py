from __future__ import division
import sys 
import win32com.client 
import visiopy


seriesdata = [100,150,200]
serieslabels = ["A","B","C"]

max_data_value = max(seriesdata)

num_datapoints = len(seriesdata)

bar_height = 1.5
bar_vsep = 0.25
max_bar_length = 4.0

dom = visiopy.DOM()
m_rect = dom.Master("rectangle", "basic_u.vss")

x = 2.0
y = 1.0

bar_rects =[]
for i in xrange(num_datapoints) :
    norm_val = seriesdata[i] / max_data_value
    cur_bar_length = norm_val * max_bar_length
    bar_rect = visiopy.Rectangle(x,y,x+cur_bar_length,y+bar_height)
    bar_rects.append(bar_rect)
    y += bar_height + bar_vsep

bar_shapes =[]

for i in xrange(num_datapoints) :
    bar_shape = dom.Drop(m_rect, bar_rects[i], serieslabels[i])
    bar_shapes.append( bar_shape)

visapp = win32com.client.Dispatch("Visio.Application") 
doc = visapp.Documents.Add("") 
page = visapp.ActivePage

with visiopy.UndoContext(visapp,"Undo1") :
    dom.Render(page)

