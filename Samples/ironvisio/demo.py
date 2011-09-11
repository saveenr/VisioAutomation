import sys
import clr
import System
import os

clr.AddReference( r"InfoGraphicsPy.Dll" )
import InfoGraphicsPy 
IG = InfoGraphicsPy 

from ironvisio import *
import charting

app = IVisio.ApplicationClass()
docs = app.Documents
doc = docs.Add("")
page = app.ActivePage

igs = IG.Session()
igs.NewDocument()

