using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Metadata.CodeGen
{
    public class VAGenCode
    {
        private VA.Metadata.MetadataDB db;

        public VAGenCode()
        {
            this.db = VA.Metadata.MetadataDB.Load();
        }

        public string GetCode()
        {

            string shape_format =
                @"
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
double |   TextBkgndTrans        |   TextBkgndTrans";


            string CONTROLCELLS =
                @"
int |   Controls_CanGlue   |   CanGlue
int |   Controls_Tip    |   Tip
double    |   Controls_X        |   X
double    |   Controls_Y         |   Y
int    |   Controls_YCon   |   YBehavior
int    |   Controls_XCon   |   XBehavior
int    |   Controls_XDyn   |   XDynamics
int    |   Controls_YDyn   |   YDynamics";


            var cg_shapeformat = create_from_text(shape_format, "ShapeFormatCells");
            var cg_controls = create_from_text(CONTROLCELLS, "ControlCells");

            var cgs = new[]
                          {
                              cg_shapeformat,
                              cg_controls
                          };

            var sb = new System.Text.StringBuilder();
            foreach (var cg in cgs)
            {
                string cg_code = cg.GenCode();
                sb.Append(cg_code);
                sb.Append("\r\n");
            }
            return sb.ToString();
        }

        public VA.Metadata.CodeGen.CellGroup create_from_text(string shape_format, string name)
        {
            var cg_shapeformat = new VA.Metadata.CodeGen.CellGroup(name);
            var lines = shape_format.Split(new[] { '\n' }).Select(s => s.Trim()).Where(s => s.Length > 0);
            foreach (string line in lines)
            {
                var tokens = line.Split(new[] { '|' }).Select(t => t.Trim()).ToArray();
                string dt = tokens[0];
                string cellname = tokens[1];
                string propname = tokens[2];

                cg_shapeformat.Add(propname, this.db.GetCellByNameCode(cellname), dt);
            }

            return cg_shapeformat;
        }

        public IEnumerable<VA.Metadata.Cell> CellsInSection(IVisio.VisSectionIndices sec)
        {
            var target_section = db.GetSectionBySectionIndex((int)sec);
            return db.CellsInSection(target_section);
        }
    }
}
