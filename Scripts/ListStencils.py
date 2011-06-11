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
    vss_files = System.IO.Directory.GetFiles(stencil_path,"*.vss")
    vst_files = System.IO.Directory.GetFiles(stencil_path,"*.vst")

    vst_dic = {}
    vss_dic = {}

    for vst_file in vst_files  :
        p,filename=os.path.split(vst_file)
        vst_name,b = os.path.splitext(filename)
        vst_dic[vst_name] = filename

    for vss_file in vss_files  :
        p,filename=os.path.split(vss_file)
        vss_name,b = os.path.splitext(filename)
        vss_dic[vss_name] = filename

    pairs = []
    for k in vss_dic.keys() :
        vss_file = vss_dic[k]
        vst_file = vst_dic.get(k,"")
        pairs.append( (vss_file,vst_file) )

    # Go through each stencil in Visio
    visapp = IVisio.ApplicationClass()
    docs = visapp.Documents
    doc = docs.Add("")
    flags= IVisio.VisOpenSaveArgs.visOpenRO | IVisio.VisOpenSaveArgs.visOpenDocked

    for vss_file, vst_file in pairs:
        stencildoc = docs.OpenEx( vss_file , flags )
        if (not print_masters) :
            print vss_file , "|", vst_file, "|", stencildoc.Title
        else:
            for master in stencildoc.Masters :
                print vss_file , "|", vst_file, "|", stencildoc.Title , "|", master.Name 
        stencildoc.Close()
    
    # Once done, close visio
    visapp.Quit()
    
main()
