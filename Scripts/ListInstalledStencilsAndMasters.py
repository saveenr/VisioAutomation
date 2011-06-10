import sys
import clr
import System
import os

clr.AddReference("Microsoft.Office.Interop.Visio")
import Microsoft.Office.Interop.Visio
IVisio = Microsoft.Office.Interop.Visio

#stencil_path = r"C:\Program Files (x86)\Microsoft Office\Office12\1033"
stencil_path = r"C:\Program Files (x86)\Microsoft Office\Office14\Visio Content\1033"

vss_files = System.IO.Directory.GetFiles(stencil_path,"*.vss")
vst_files = System.IO.Directory.GetFiles(stencil_path,"*.vst")

vst_dic = {}
vss_dic = {}
for vst_file in vst_files  :
    p,n=os.path.split(vst_file)
    a,b = os.path.splitext(n)
    vst_dic[a] = n

for vss_file in vss_files  :
    p,n=os.path.split(vss_file)
    a,b = os.path.splitext(n)
    vss_dic[a] = n

pairs = []
for k in vss_dic.keys() :
    a = vss_dic[k]
    b = vst_dic.get(k,None)
    pairs.append( (a,b) )
    
visapp = IVisio.ApplicationClass()

doc = visapp.Documents.Add("")
for a,b in pairs:
    flags= IVisio.VisOpenSaveArgs.visOpenRO | IVisio.VisOpenSaveArgs.visOpenDocked
    stencildoc = visapp.Documents.OpenEx( a , flags )
    print a, "|", stencildoc.Title, "|", stencildoc.Subject
    stencildoc.Close()
    #stencildoc = visapp.Documents.OpenEx( b , flags )
    #print stencildoc.Fullname , ",", stencildoc.Title
    #stencildoc.Close()
    #for master in stencildoc.Masters :
    #    print master.Name, 
