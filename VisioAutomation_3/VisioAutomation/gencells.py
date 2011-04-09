import sys
import os

# generate common ShapeSheet objects

#type, cellsrc, name
XFORMCELLS = """
double    |    PinX  	   |      PinX  	  
double    |    PinY  	   |      PinY  	  
double    |    LocPinX	   |      LocPinX	  
double    |    LocPinY	   |      LocPinY	  
double    |    Width 	   |      Width 	  
double    |    Height 	   |      Height 	  
double    |    Angle 	   |      Angle 	  
"""

CONTROLCELLS = """
int |   Controls_CanGlue   |   CanGlue
int |   Controls_Tip    |   Tip
double    |   Controls_X 	   |   X
double    |   Controls_Y  	   |   Y
int    |   Controls_YCon   |   YBehavior
int    |   Controls_XCon   |   XBehavior
int    |   Controls_XDyn   |   XDynamics
int    |   Controls_YDyn   |   YDynamics
"""

LOCKCELLS = """
bool   |  LockAspect			 |    LockAspect
bool   |  LockBegin 			 |    LockBegin
bool   |  LockCalcWH			 |    LockCalcWH
bool   |  LockCrop 				 |    LockCrop
bool   |  LockCustProp			 |    LockCustProp
bool   |  LockDelete			 |    LockDelete
bool   |  LockEnd 				 |    LockEnd
bool   |  LockFormat			 |    LockFormat
bool   |  LockFromGroupFormat	 |    LockFromGroupFormat
bool   |  LockGroup 			 |    LockGroup
bool   |  LockHeight			 |    LockHeight
bool   |  LockMoveX 			 |    LockMoveX
bool   |  LockMoveY 			 |    LockMoveY
bool   |  LockRotate			 |    LockRotate
bool   |  LockSelect			 |    LockSelect
bool   |  LockTextEdit			 |    LockTextEdit
bool   |  LockThemeColors 		 |    LockThemeColors
bool   |  LockThemeEffects		 |    LockThemeEffects
bool   |  LockVtxEdit			 |    LockVtxEdit
bool   |  LockWidth 			 |    LockWidth
"""


SHAPEFORMAT="""
int    |   FillBkgnd             |   FillBkgnd
double |   FillBkgndTrans        |   FillBkgndTrans
int    |   FillForegnd           |   FillForegnd
double |   FillForegndTrans      |   FillForegndTrans
int    |   FillPattern           |   FillPattern
double |   ShapeShdwObliqueAngle |   ShapeShdwObliqueAngle
double |   ShapeShdwOffsetX      |   ShapeShdwOffsetX
double |   ShapeShdwOffsetY      |   ShapeShdwOffsetY
double |   ShapeShdwScaleFactor  |   ShapeShdwScaleFactor
int    |   ShapeShdwType         |   ShapeShdwType
int    |   ShdwBkgnd             |   ShdwBkgnd
double |   ShdwBkgndTrans        |   ShdwBkgndTrans
int    |   ShdwForegnd           |   ShdwForegnd
double |   ShdwForegndTrans      |   ShdwForegndTrans
int    |   ShdwPattern           |   ShdwPattern
int    |   BeginArrow            |   BeginArrow
double |   BeginArrowSize        |   BeginArrowSize
int    |   EndArrow              |   EndArrow
double |   EndArrowSize          |   EndArrowSize
int    |   LineCap               |   LineCap
int    |   LineColor             |   LineColor
double |   LineColorTrans        |   LineColorTrans
int    |   LinePattern           |   LinePattern
double |   LineWeight            |   LineWeight
double |   Rounding              |   Rounding
int    |   Char_Font              |   CharFont
int    |   Char_Color             |   CharColor
double |   Char_ColorTrans        |   CharColorTrans
double |   Char_Size              |   CharSize
int    |   TextBkgnd             |   TextBkgnd
double |   TextBkgndTrans        |   TextBkgndTrans
"""
def printtop() :
    print """
using System;
using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
    """
    
def gencode_for_cells(text,classname,queryname,qt,si) :
    lines = text.strip()
    lines = text.split("\n");
    lines = [l for l in lines if len(l)]


    printtop()

    if (qt=="Cell") :
        baseclassname = "VA.ShapeSheet.DataGroup"
    else :
        baseclassname = "VA.ShapeSheet.CellSectionDataGroup"

    print "----------------------------------"
    print "public partial class " + classname + " : " + baseclassname
    print "{"


    data = []
    for line in lines :
        tokens = [ t.strip() for t in line.split("|")]
        celltype = tokens[0]
        cellsrc = tokens[1]
        cellname = tokens[2]
        data.append((celltype,cellsrc,cellname))
        
        print "    public VA.ShapeSheet.CellData<"+celltype+"> " + cellname + " { get; set; }"



    print
    if (qt=="Cell") :
        print "    internal void _Apply( System.Action<VA.ShapeSheet.SRC,VA.ShapeSheet.FormulaLiteral> func)"
    elif (qt=="Section") :
        print "    internal void _Apply( System.Action<VA.ShapeSheet.SRC,VA.ShapeSheet.FormulaLiteral> func, short row)"
    print "    {"
    for celltype,cellsrc,cellname in data:
        if (qt=="Cell") :
            print"            func(ShapeSheet.SRCConstants.", cellsrc , " , this." , cellname , ".Formula);"
        elif (qt=="Section") :
            print"            func(VA.ShapeSheet.SRCConstants.", cellsrc , ".ForRow(row) , this." , cellname , ".Formula);"

    print "    }"

    print
    print "   private static " + classname + " get_cells_from_row","(" , queryname, "query,VA.ShapeSheet.Query.QueryDataSet<double> qds, int row)"
    print "   {"
    print "      var cells = new ", classname,"();"
    for celltype,cellsrc,cellname in data:
        x = ""
        if ( celltype=="int") : x = ",v => (int)v"
        elif ( celltype=="bool") : x = ",v => (bool)v"
        print "      cells.", cellname, "= qds.GetItem(row, query." ,cellname, x ,");"
    print "      return cells;"
    print "   }"


    print
    print "    internal static IList<", classname , "> GetCells(IVisio.Page page, IList<int> shapeids)"
    print "    {"
    print "      var query = new ", queryname,"();"
    print "      return " + baseclassname + "._GetCells(page, shapeids, query, get_cells_from_row);"
    print "    }"
    print

    print "    internal static ", classname , " GetCells(IVisio.Shape shape)"
    print "    {"
    print "      var query = new ", queryname,"();"
    print "      return " + baseclassname + "._GetCells(shape, query, get_cells_from_row);"
    print "    }"
    print

    print "}"    

    print "----------------------------------"
    printtop()

    print
    print "public class", queryname, ": VA.ShapeSheet.Query." + qt + "Query"
    print "{"
    for celltype,cellsrc,cellname in data:
        print"   public VA.ShapeSheet.Query." + qt + "QueryColumn", cellname , " {get; set;}"

    print
    print "    public ",queryname,"() :"
    print "            base(IVisio.VisSectionIndices.",si,")"
    print "    {"
    for celltype,cellsrc,cellname in data:
        if (qt=="Cell"):
            print "        this."+ cellname+" = this.AddColumn(VA.ShapeSheet.SRCConstants.", cellsrc,", \""+cellname+"\");"
        elif (qt=="Section"):
            print "        this."+ cellname+" = this.AddColumn(VA.ShapeSheet.SRCConstants.", cellsrc,".Cell, \""+cellname+"\");"
    print "    }"

    print 
    print "}"    

    print "----------------------------------"
    printtop()

    print
    print "public static class", classname+"Helper"
    print "{"
    for celltype,cellsrc,cellname in data:
        print"            public VA.ShapeSheet.Query." + qt + "QueryColumn", cellname , " {get; set;}"




#gencode_for_cells(XFORMCELLS, "XFormCells", "XFormQuery","Cell","")
gencode_for_cells(CONTROLCELLS, "ControlCells", "ControlQuery","Section","visSectionControls")
#gencode_for_cells(LOCKCELLS, "LockCells", "LockQuery","Cell","")
#gencode_for_cells(SHAPEFORMAT, "ShapeFormatCells", "ShapeFormatQuery","Cell","")

    
