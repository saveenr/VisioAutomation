import sys
import os
import clr

script_path = os.path.dirname(__file__)
print script_path

# Load Visio Typelib
# Loading Visio typelib
visio_asm = clr.AddReference("Microsoft.Office.Interop.Visio")
import Microsoft.Office.Interop.Visio

# Load VisioAutomation
visauto_path = script_path
visauto_assemblies = [
    "VisioAutomation.dll",
    "VisioAutomation.DocumentAnalysis.dll",
    "VisioAutomation.Models.dll",
    "VisioAutomation.Scripting.dll",
    ]

for asm in visauto_assemblies :
    print "Loading", asm
    clr.AddReferenceToFileAndPath( os.path.join(visauto_path, asm ) )

import VisioAutomation
import VisioAutomation.DocumentAnalysis
import VisioAutomation.Models
import VisioAutomation.Scripting

context = VisioAutomation.Scripting.DefaultContext()
client = VisioAutomation.Scripting.Client(None,context)
