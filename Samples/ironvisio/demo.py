import sys
import clr
import System
import os

script_path = System.IO.Path.GetDirectoryName(__file__)
visutildll_file = System.IO.Path.Combine( script_path, r"VisUtil/bin/Debug/VisUtil.Dll" )
clr.AddReferenceToFileAndPath( visutildll_file )
import VisUtil


from ironvisio import *
import charting

app = IVisio.ApplicationClass()
docs = app.Documents
doc = docs.Add("")
page = app.ActivePage

VisUtil.DrawUtil.DrawCircleFromCenter( page, 2, 4, 1.5)

System.Console.ReadKey()
