import sys 
import win32com.client 
import visiopy

visapp = win32com.client.Dispatch("Visio.Application") 
   
doc = visapp.Documents.Add("") 
page = visapp.ActivePage

dom = visiopy.DOM()
s0 = dom.Drop("rectangle", "basic_u.vss", visiopy.Rectangle(0,0,1,1), "A")
s1 = dom.Drop("octagon", "basic_u.vss", visiopy.Point(4,3), "B")
c0 = dom.Drop("dynamic connector", "basic_u.vss", visiopy.Point(-1,-1), "C")
dom.Connect(s0,s1,c0)

dom.Render(page)

q = visiopy.Query()
q.Add( dom.Shapes[0].VisioShapeID, visiopy.SRCConstants.Width )
q.Add( dom.Shapes[0].VisioShapeID, visiopy.SRCConstants.Height )
formulas = q.GetFormulas(page)
results = q.GetResults(page)
print formulas
print results

#u = visiopy.Update()
#u.Add( dom.Shapes[0].VisioShapeID, visiopy.SRCConstants.Width , "5")
#u.Add( dom.Shapes[0].VisioShapeID, visiopy.SRCConstants.Height , "3")
#result = u.SetFormulas(page)

print result

