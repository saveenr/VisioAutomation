import sys
import clr
import System
import os

script_path = System.IO.Path.GetDirectoryName(__file__)
clr.AddReferenceToFileAndPath( System.IO.Path.Combine( script_path, r"InfoGraphicsPy/bin/Debug/InfoGraphicsPy.Dll" ) )
import InfoGraphicsPy 
IG = InfoGraphicsPy 

from ironvisio import *
import charting

app = IVisio.ApplicationClass()
docs = app.Documents
doc = docs.Add("")
page = app.ActivePage

center = IG.Point(2,4)
IG.DrawUtil.DrawCircleFromCenter( page, center, 1.5)

System.Console.ReadKey()

