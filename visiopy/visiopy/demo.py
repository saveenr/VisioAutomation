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
dom.Drop(masterrect, visiopy.Point(1,1))
dom.Drop(masteroctagon, visiopy.Point(4,3))
dom.Drop(masterconnector, visiopy.Point(-1,-1))
dom.Render()

q = visiopy.Query()
q.Add( shapes[0].ID16, visiopy.SRCConstants.Width )
q.Add( shapes[0].ID16, visiopy.SRCConstants.Height )
formulas = q.GetFormulas(page)
results = q.GetResults(page)
print formulas
print results

u = visiopy.Update()
u.Add( shapes[0].ID16, visiopy.SRCConstants.Width , "5")
u.Add( shapes[0].ID16, visiopy.SRCConstants.Height , "3")
result = u.SetFormulas(page)

print result

visiopy.connect(shapes[0], shapes[1], shapes[2])
