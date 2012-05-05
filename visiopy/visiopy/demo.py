from __future__ import division
import sys 
import win32com.client 
import visiopy

visapp = win32com.client.Dispatch("Visio.Application") 
   
with visiopy.UndoContext(visapp,"Undo1") :
    doc = visapp.Documents.Add("") 
    page = visapp.ActivePage

    default_fmt = { visiopy.SRCConstants.FillForegnd : "rgb(251,55,1)" }

    dom = visiopy.DOM()
    m_rect = dom.Master("rectangle", "basic_u.vss")
    m_octogon = dom.Master("octagon", "basic_u.vss")
    m_dyncon= dom.Master("dynamic connector", "basic_u.vss")

    s0 = dom.Drop(m_rect, visiopy.Rectangle(0,0,1,1), "A", default_fmt)
    s1 = dom.Drop(m_octogon, visiopy.Point(4,3), "B", default_fmt)

    dom.Connect(s0,s1,m_dyncon)
    dom.Render(page)

    srcs = [visiopy.SRCConstants.Width , visiopy.SRCConstants.Height,visiopy.SRCConstants.PinX, visiopy.SRCConstants.PinY]
    shapeids = [dom.Shapes[0].VisioShapeID,dom.Shapes[1].VisioShapeID]

    formulas = visiopy.Query.QueryFormulas(page,shapeids, srcs)
    results = visiopy.Query.QueryResults(page,shapeids, srcs)
    print formulas
    print results

