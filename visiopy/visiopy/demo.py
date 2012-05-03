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
s1 = dom.Drop(m_octogon, visiopy.Point(4,3), "B")
c0 = dom.Drop(m_dyncon, visiopy.Point(-1,-1), "C")

dom.Connect(s0,s1,c0)

dom.Render(page)

srcs = [visiopy.SRCConstants.Width , visiopy.SRCConstants.Height ]
shapeids = [dom.Shapes[0].VisioShapeID]

print dir(visiopy)

q = visiopy.QueryEx(shapeids, srcs)
formulas = q.GetFormulas(page)
results = q.GetResults(page)
print formulas
print results

