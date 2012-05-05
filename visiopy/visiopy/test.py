from __future__ import division
import sys 
import win32com.client 
import visiopy

visapp = win32com.client.Dispatch("Visio.Application") 

with visiopy.UndoContext(visapp,"Undo1") :
    doc = visapp.Documents.Add("") 
    page = visapp.ActivePage

    dom = visiopy.DOM()
    m_rect = dom.Master("rectangle", "basic_u.vss")
    m_octogon = dom.Master("octagon", "basic_u.vss")
    m_dyncon= dom.Master("dynamic connector", "basic_u.vss")

    s0 = dom.Drop(m_rect, visiopy.Rectangle(0,0,1,1), "A")
    s1 = dom.Drop(m_octogon, visiopy.Point(4,3), "B",)

    dom.Connect(s0,s1,m_dyncon,"C0")
    dom.Connect(s0,s1,m_dyncon,"C1")
    dom.Connect(s1,s1,m_dyncon,"C2", cells={visiopy.SRCConstants.LineColor: "rgb(255,128,64)"} )

    dom.Render(page)

    srcs = [visiopy.SRCConstants.Width , visiopy.SRCConstants.Height,visiopy.SRCConstants.PinX, visiopy.SRCConstants.PinY]
    shapeids = [dom.Shapes[0].VisioShapeID,dom.Shapes[1].VisioShapeID]
