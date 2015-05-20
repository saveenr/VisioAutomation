import sys
import clr
import System
import os

clr.AddReference("Microsoft.Office.Interop.Visio")
import Microsoft.Office.Interop.Visio
IVisio = Microsoft.Office.Interop.Visio

lang_to_id = {
              "en" : "1033"
              }

ver_to_path = {
               "2007" : r"C:\Program Files (x86)\Microsoft Office\Office12",
               "2010" : r"C:\Program Files (x86)\Microsoft Office\Office14\Visio Content"
               }

def main() :
    visio_version = "2010"
    print_masters = True

    stencil_path = System.IO.Path.Combine( ver_to_path[visio_version] , lang_to_id["en"] )
    vst_files = System.IO.Directory.GetFiles(stencil_path,"*.vst")

    visapp = IVisio.ApplicationClass()
    docs = visapp.Documents
    flags= IVisio.VisOpenSaveArgs.visOpenRO | IVisio.VisOpenSaveArgs.visOpenDocked

    for vst_file in vst_files:
        
        doc = docs.Open( System.IO.Path.Combine( stencil_path, vst_file) )
        print
        print "--------------"
        print doc.Name, doc.Title
        n=0
        for d in docs :
            if n>0:
                print d.Name, d.Title
            n+=1
        doc.Close()
    
    # Once done, close visio
    visapp.Quit()
    
main()
