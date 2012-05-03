import sys 
import win32com.client 
import visiopy

visapp = win32com.client.Dispatch("Visio.Application") 
   
doc = visapp.Documents.Add("") 
page = visapp.ActivePage

dom = visiopy.DOM()
m_rect = dom.Master("rectangle", "basic_u.vss")
m_octogon = dom.Master("octagon", "basic_u.vss")
m_dyncon= dom.Master("dynamic connector", "basic_u.vss")

s0 = dom.Drop(m_rect, visiopy.Rectangle(0,0,1,1), "A")
s1 = dom.Drop(m_octogon, visiopy.Point(4,3), "B",)
c0 = dom.Drop(m_dyncon, visiopy.Point(-1,-1), "C")
c1 = dom.Drop(m_dyncon, visiopy.Point(-1,-1), "D")
c2 = dom.Drop(m_dyncon, visiopy.Point(-1,-1), "E")

dom.AutoConnect(s0,s1,c0)
dom.AutoConnect(s0,s1,c1)
dom.AutoConnect(s1,s1,c2)

dom.Render(page)

srcs = [visiopy.SRCConstants.Width , visiopy.SRCConstants.Height,visiopy.SRCConstants.PinX, visiopy.SRCConstants.PinY]
shapeids = [dom.Shapes[0].VisioShapeID,dom.Shapes[1].VisioShapeID]
