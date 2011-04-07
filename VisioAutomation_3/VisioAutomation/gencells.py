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

    print "----------------------------------"
    print "public class", classname
    print "{"

    data = []
    for line in lines :
        tokens = [ t.strip() for t in line.split("|")]
        celltype = tokens[0]
        cellsrc = tokens[1]
        cellname = tokens[2]
        data.append((celltype,cellsrc,cellname))
        
        print "public VA.ShapeSheet.CellData<", celltype, ">" , cellname, "{ get; set; }"

    if (qt=="Cell") :
        print """
                public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short id)
                {
                    this._Apply((src, f) => update.SetFormulaIgnoreNull(id, src, f));
                }

                public void Apply(VA.ShapeSheet.Update.SRCUpdate update)
                {
                    this._Apply((src, f) => update.SetFormulaIgnoreNull(src, f));
                }
        """
    elif (qt=="Section") :
        
        print """
            public void Apply(VA.ShapeSheet.Update.SRCUpdate update, short row)
            {
                this._Apply((src, f) => update.SetFormulaIgnoreNull(src, f),row);
            }

            public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short id, short row )
            {
                this._Apply((src, f) => update.SetFormulaIgnoreNull(id, src, f), row);
            }
        """




    if (qt=="Cell") :
        print "internal void _Apply( System.Action<VA.ShapeSheet.SRC,VA.ShapeSheet.FormulaLiteral> func)"
    elif (qt=="Section") :
        print "internal void _Apply( System.Action<VA.ShapeSheet.SRC,VA.ShapeSheet.FormulaLiteral> func, short row)"


    print "{"

    for celltype,cellsrc,cellname in data:
        if (qt=="Cell") :
            print"            func(ShapeSheet.SRCConstants.", cellsrc , " , this." , cellname , ".Formula);"
        elif (qt=="Section") :
            print"            func(VA.ShapeSheet.SRCConstants.", cellsrc , ".ForRow(row) , this." , cellname , ".Formula);"

    print "}"    

    print "}"    

    print "----------------------------------"
    printtop()

    print
    print "public class", queryname, ": VA.ShapeSheet.Query." + qt + "Query"
    print "{"
    for celltype,cellsrc,cellname in data:
        print"            public VA.ShapeSheet.Query." + qt + "QueryColumn", cellname , " {get; set;}"

    print
    print "public ",queryname,"() :"
    print "            base(IVisio.VisSectionIndices.",si,")"
    print "{"    
    for celltype,cellsrc,cellname in data:
        if (qt=="Cell"):
            print "    this.", cellname," = this.AddColumn(VA.ShapeSheet.SRCConstants.", cellsrc,", \""+cellname+"\");"
        elif (qt=="Section"):
            print "    this.", cellname," = this.AddColumn(VA.ShapeSheet.SRCConstants.", cellsrc,".Cell, \""+cellname+"\");"
    print "}"    

    print 
    print "}"    

    print "----------------------------------"
    printtop()

    print
    print "public static class", classname+"Helper"
    print "{"
    for celltype,cellsrc,cellname in data:
        print"            public VA.ShapeSheet.Query." + qt + "QueryColumn", cellname , " {get; set;}"

    print
    print "private static ",classname, " get_" + classname + "_from_row","(" , queryname, "query,VA.ShapeSheet.Query.QueryDataSet<double> qds, int row)"
    print "{"
    print "  var cells = new ", classname,"();"
    for celltype,cellsrc,cellname in data:
        x = ""
        if ( celltype=="int") : x = ",v => (int)v"
        elif ( celltype=="bool") : x = ",v => (bool)v"
        print "   cells.", cellname, "= qds.GetItem(row, query." ,cellname, x ,");"
    print "return cells;"
    print "}"

    if (qt=="Cell") :
        print
        print "public static IList<", classname , "> GetCells(IVisio.Page page, IList<int> shapeids)"
        print "{"
        print "  var query = new ", queryname,"();"
        print "  var qds = query.GetFormulasAndResults<double>(page, shapeids);"
        print "  var cells_list = new List<", classname,">(shapeids.Count);"
        print "  for (int i = 0; i < qds.RowCount; i++)"
        print "  {"
        print "     var cells = get_" + classname + "_from_row(query, qds, i);"
        print "     cells_list.Add( cells );"
        print "  }"
        print ""
        print "  return cells_list;"
        print "}"
        print

        print
        print "public static ", classname , " GetCells(IVisio.Shape shape)"
        print "{"
        print "  var query = new ", queryname,"();"
        print "  var qds = query.GetFormulasAndResults<double>(shape);"
        print "  var cells = get_" + classname + "_from_row(query, qds, 0);"
        print "  return cells;"
        print "}"
        print

        print
        print "}"

    if (qt=="Section") :
        print
        print "public static IList<List<", classname , ">> GetCells(IVisio.Page page, IList<int> shapeids)"
        print "{"
        print "  var query = new ", queryname,"();"
        print "  var qds = query.GetFormulasAndResults<double>(page, shapeids);"
        print "  var list = new List<List<ControlCells>>(shapeids.Count);"
        print "  foreach (var group in qds.Groups)"
        print "  {"
        print "     var cells_list = new List<",classname,">(group.Count);"
        print "     if (group.Count>0)"
        print "     {"
        print "        for (int i = 0; i < qds.RowCount; i++)"
        print "        {"
        print "           var cells = get_" + classname + "_from_row(query, qds, i);"
        print "           cells_list.Add( cells );"
        print "        }"
        print "     }"
        print "     list.Add( cells_list );"
        print "  }"
        print "  return list;"
        print "}"
        print

        print
        print "public static IList<", classname , "> GetCells(IVisio.Shape shape)"
        print "{"
        print "  var query = new ", queryname,"();"
        print "  var qds = query.GetFormulasAndResults<double>(shape);"
        print "  var cells_list = new List<",classname,">(qds.RowCount);"
        print "  for (int row = 0; row < qds.RowCount; row++)"
        print "  {"
        print "      var cells = get_" + classname + "_from_row(query, qds, row);"
        print "      cells_list.Add(cells);"
        print "  }"
        print "  return cells_list;"
        print "}"
        print

        print
        print "}"

#gencode_for_cells(XFORMCELLS, "XFormCells", "XFormQuery","Cell","")
gencode_for_cells(CONTROLCELLS, "ControlCells", "ControlQuery","Section","visSectionControls")
#gencode_for_cells(LOCKCELLS, "LockCells", "LockQuery","Cell","")

    
