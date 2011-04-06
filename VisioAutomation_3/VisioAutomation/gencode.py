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
double |   Controls_Glue   |   Glue
double |   Controls_Tip    |   Tip 
int    |   Controls_X 	   |   X 
int    |   Controls_Y  	   |   Y  
int    |   Controls_YCon   |   YCon
int    |   Controls_XCon   |   XCon
int    |   Controls_XDyn   |   XDyn
int    |   Controls_YDyn   |   YDyn  
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
    
def x(text,classname,queryname,qt,si) :
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
                public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short id, short row)
                {
                    this._Apply((src, f) => update.SetFormulaIgnoreNull(id, src, f));
                }

                public void Apply(VA.ShapeSheet.Update.SRCUpdate update, short row)
                {
                    this._Apply((src, f) => update.SetFormulaIgnoreNull(src, f));
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
            print"            func(ShapeSheet.SRCConstants.", cellsrc , ".Cell , this." , cellname , ".Formula, row);"

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

x(XFORMCELLS, "XFormCells", "XFormQuery","Cell","")
x(CONTROLCELLS, "ControlCells", "ControlQuery","Section","visSectionControls")
    
    
