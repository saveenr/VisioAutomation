import sys
import os
import clr

script_path = os.path.dirname(os.path.realpath(__file__))

if (script_path==None) :
    raise Exception("Script path could not be determined")


# Load Visio Typelib
# Loading Visio typelib
visio_asm = clr.AddReference("Microsoft.Office.Interop.Visio")
import Microsoft.Office.Interop.Visio

# Load VisioAutomation
visauto_nuget_package_name = "VisioAutomation2010"
visauto_path = os.path.join(script_path , "packages", visauto_nuget_package_name )
visauto_dll_path = None
visauto_dll_path_nuget = os.path.join(visauto_path, "lib", "net40" )
visauto_dll_path_bindebug = os.path.join( os.path.dirname(script_path), "VisioAutomation.Scripting", "bin", "Debug" )

trypaths = [ ("NuGet",visauto_dll_path_nuget), ("LocalCompile",visauto_dll_path_bindebug)]
for name,path in trypaths :
    if (os.path.exists(path)) :
        print "Loading VisioAutomation assemblies from ", path
        visauto_dll_path = path
if (visauto_dll_path == None) :
    raise Exception("Could not find either nuget binaries or local binaries for VisioAutomation.Scripting")            

visauto_assemblies = [
    "VisioAutomation.dll",
    "VisioAutomation.DocumentAnalysis.dll",
    "VisioAutomation.Models.dll",
    "VisioAutomation.Scripting.dll",
    ]

for asm in visauto_assemblies :
    print "Loading", asm
    asm_full_path = os.path.join(visauto_dll_path, asm )
    clr.AddReferenceToFileAndPath( asm_full_path )

import VisioAutomation
import VisioAutomation.DocumentAnalysis
import VisioAutomation.Models
import VisioAutomation.Scripting

context = VisioAutomation.Scripting.DefaultContext()
client = VisioAutomation.Scripting.Client(None,context)
client.VerboseLogging = Falses