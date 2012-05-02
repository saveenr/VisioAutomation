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

dropdata = [(masterrect, (1,1) ),
    (masteroctagon, (4,3) ),
    (masterconnector, (-1,-1) )
    ]

shapes = []
for dd in dropdata:
    master = dd[0]
    shape = page.Drop( master, *dd[1]) 
    shapes.append(shape)

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
