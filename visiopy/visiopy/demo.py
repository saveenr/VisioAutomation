import sys 
import win32com.client 
import visiopy

visapp = win32com.client.Dispatch("Visio.Application") 
   
doc = visapp.Documents.Add("") 
page = visapp.ActivePage

stencilname = "basic_u.vss" 
stencildoc = visiopy.openstencil(visapp.Documents,stencilname)

masterrect = stencildoc.Masters.ItemU("rectangle") 
masteroctagon = stencildoc.Masters.ItemU("octagon") 
masterconnector= stencildoc.Masters.ItemU("dynamic connector")

dom = visiopy.DOM()
s0 = dom.Drop(masterrect, visiopy.Point(1,1))
s1 = dom.Drop(masteroctagon, visiopy.Point(4,3))
c0 = dom.Drop(masterconnector, visiopy.Point(-1,-1))
dom.Connect(s0,s1,c0)
dom.Render(page)

q = visiopy.Query()
q.Add( dom.Shapes[0].VisioShapeID, visiopy.SRCConstants.Width )
q.Add( dom.Shapes[0].VisioShapeID, visiopy.SRCConstants.Height )
formulas = q.GetFormulas(page)
results = q.GetResults(page)
print formulas
print results

u = visiopy.Update()
u.Add( dom.Shapes[0].VisioShapeID, visiopy.SRCConstants.Width , "5")
u.Add( dom.Shapes[0].VisioShapeID, visiopy.SRCConstants.Height , "3")
result = u.SetFormulas(page)

print result

