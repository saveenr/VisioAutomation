from __future__ import division
import sys 
import win32com.client 
import visiopy


def get_bar_chart() :
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

    return dom

def get_column_chart() :

    #6000	13000	35000
    seriesdata = [6,13,35]
    serieslabels = [ str(d)+"K" for d in seriesdata ]
    seriesbaselabels = ["2010","2011","2012"]

    max_data_value = max(seriesdata)

    num_datapoints = len(seriesdata)

    bar_width = 1.5
    bar_hsep = 0.25
    max_bar_height = 4.0

    dom = visiopy.DOM()
    m_rect = dom.Master("rectangle", "basic_u.vss")

    x = 2.0
    y = 1.0

    label_height = 0.25

    bar_rects =[]
    label_rects =[]
    for i in xrange(num_datapoints) :
        norm_val = seriesdata[i] / max_data_value
        cur_bar_width = norm_val * max_bar_height
        bar_rect = visiopy.Rectangle(x,y,x+bar_width,y+cur_bar_width)
        bar_rects.append(bar_rect)
        label_rect = visiopy.Rectangle( bar_rect.Left, bar_rect.Bottom - label_height, bar_rect.Right, bar_rect.Bottom )
        label_rects.append(label_rect)        
        x += bar_width + bar_hsep

    bar_shapes =[]
    label_shapes = []

    for i in xrange(num_datapoints) :
        bar_shape = dom.Drop(m_rect, bar_rects[i], serieslabels[i])
        bar_shapes.append( bar_shape)

    for i in xrange(num_datapoints) :
        label_shape = dom.Drop(m_rect, label_rects[i], seriesbaselabels[i])
        label_shapes.append( label_shape )

    return dom



visapp = win32com.client.Dispatch("Visio.Application") 
doc = visapp.Documents.Add("") 
page = visapp.ActivePage

dom = get_column_chart()

with visiopy.UndoContext(visapp,"Undo1") :
    dom.Render(page)

