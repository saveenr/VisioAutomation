using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Update;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeGeometry
{
    public class GeometrySection
    {
        public List<GeometryRow> Rows { get; private set; }
        public VA.ShapeSheet.FormulaLiteral NoFill { get; set; }
        public VA.ShapeSheet.FormulaLiteral NoLine { get; set; }
        public VA.ShapeSheet.FormulaLiteral NoShow { get; set; }
        public VA.ShapeSheet.FormulaLiteral NoSnap { get; set; }

        public GeometrySection()
        {
            this.Rows = new List<GeometryRow>();
        }
        
        public short Render(IVisio.Shape shape)
        {
            short sec_index = ShapeGeometryHelper.AddGeometrySection(shape);
            short row_count = shape.RowCount[sec_index];

            var update = new VA.ShapeSheet.Update.SRCUpdate();

            var src_nofill = VA.ShapeSheet.SRCConstants.Geometry_NoFill.ForSectionAndRow(sec_index, 0);
            var src_noline = VA.ShapeSheet.SRCConstants.Geometry_NoLine.ForSectionAndRow(sec_index, 0);
            var src_noshow = VA.ShapeSheet.SRCConstants.Geometry_NoShow.ForSectionAndRow(sec_index, 0);
            var src_nosnap = VA.ShapeSheet.SRCConstants.Geometry_NoSnap.ForSectionAndRow(sec_index, 0);

            update.SetFormulaIgnoreNull(src_nofill, this.NoFill);
            update.SetFormulaIgnoreNull(src_noline, this.NoLine);
            update.SetFormulaIgnoreNull(src_noshow, this.NoShow);
            update.SetFormulaIgnoreNull(src_nosnap, this.NoSnap);

            foreach (var row in this.Rows)
            {
                if (row is RowMoveTo)
                {
                    var moveto_row = (RowMoveTo) row;
                    CreateMoveToRow(moveto_row, shape, update, row_count, sec_index);
                }
                else if (row is RowLineTo)
                {
                    var lineto_row = (RowLineTo)row;
                    CreateLineToRow(lineto_row, shape, update, row_count, sec_index);
                }
                else if (row is RowArcTo)
                {
                    var arcto_row = (RowArcTo)row;
                    CreateArcToRow(arcto_row, shape, update, row_count, sec_index);
                }
                else if (row is RowEllipticalArcTo)
                {
                    var ellipticalarcto_row = (RowEllipticalArcTo)row;
                    CreateEllipticalArcToRow(ellipticalarcto_row, shape, update, row_count, sec_index);
                }
                else
                {
                    string msg = string.Format("Unsupported row type \"{0}\"", row.GetType().Name);
                    throw new AutomationException(msg);
                }
                row_count++;
            }

            update.Execute(shape);
            return 0;
        }

        private static void CreateEllipticalArcToRow(RowEllipticalArcTo rowdef, Shape shape, SRCUpdate update, short row, short section)
        {
            short row_index = shape.AddRow(section, row, (short) IVisio.VisRowTags.visTagArcTo);
            var x_src = VA.ShapeSheet.SRCConstants.Geometry_X.ForSectionAndRow(section, row_index);
            var y_src = VA.ShapeSheet.SRCConstants.Geometry_Y.ForSectionAndRow(section, row_index);
            var a_src = VA.ShapeSheet.SRCConstants.Geometry_A.ForSectionAndRow(section, row_index);
            var b_src = VA.ShapeSheet.SRCConstants.Geometry_B.ForSectionAndRow(section, row_index);
            var c_src = VA.ShapeSheet.SRCConstants.Geometry_C.ForSectionAndRow(section, row_index);
            var d_src = VA.ShapeSheet.SRCConstants.Geometry_D.ForSectionAndRow(section, row_index);
            update.SetFormulaIgnoreNull(x_src, rowdef.X);
            update.SetFormulaIgnoreNull(y_src, rowdef.Y);
            update.SetFormulaIgnoreNull(a_src, rowdef.A);
            update.SetFormulaIgnoreNull(b_src, rowdef.B);
            update.SetFormulaIgnoreNull(c_src, rowdef.C);
            update.SetFormulaIgnoreNull(d_src, rowdef.D);
        }

        private static void CreateArcToRow(RowArcTo rowdef, Shape shape, SRCUpdate update, short row, short section)
        {
            short row_index = shape.AddRow(section, row, (short) IVisio.VisRowTags.visTagArcTo);
            var x_src = VA.ShapeSheet.SRCConstants.Geometry_X.ForSectionAndRow(section, row_index);
            var y_src = VA.ShapeSheet.SRCConstants.Geometry_Y.ForSectionAndRow(section, row_index);
            var a_src = VA.ShapeSheet.SRCConstants.Geometry_A.ForSectionAndRow(section, row_index);
            update.SetFormulaIgnoreNull(x_src, rowdef.X);
            update.SetFormulaIgnoreNull(y_src, rowdef.Y);
            update.SetFormulaIgnoreNull(a_src, rowdef.A);
        }

        private static void CreateLineToRow(RowLineTo rowdef, Shape shape, SRCUpdate update, short row, short section)
        {
            short row_index = shape.AddRow(section, row, (short) IVisio.VisRowTags.visTagLineTo);
            var x_src = VA.ShapeSheet.SRCConstants.Geometry_X.ForSectionAndRow(section, row_index);
            var y_src = VA.ShapeSheet.SRCConstants.Geometry_Y.ForSectionAndRow(section, row_index);
            update.SetFormulaIgnoreNull(x_src, rowdef.X);
            update.SetFormulaIgnoreNull(y_src, rowdef.Y);
        }

        private static void CreateMoveToRow(RowMoveTo rowdef, Shape shape, SRCUpdate update, short row, short section)
        {
            short row_index = shape.AddRow(section, row, (short) IVisio.VisRowTags.visTagMoveTo);
            var x_src = VA.ShapeSheet.SRCConstants.Geometry_X.ForSectionAndRow(section, row_index);
            var y_src = VA.ShapeSheet.SRCConstants.Geometry_Y.ForSectionAndRow(section, row_index);
            update.SetFormulaIgnoreNull(x_src, rowdef.X);
            update.SetFormulaIgnoreNull(y_src, rowdef.Y);
        }

        public void MoveTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y)
        {
            var row = new VA.ShapeGeometry.RowMoveTo(x, y);
            this.Rows.Add(row);
        }

        public void LineTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y)
        {
            var row = new VA.ShapeGeometry.RowLineTo(x, y);
            this.Rows.Add(row);
        }

        public void ArcTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a)
        {
            var row = new VA.ShapeGeometry.RowArcTo(x, y, a );
            this.Rows.Add(row);
        }

        public void EllipticalArcTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a, VA.ShapeSheet.FormulaLiteral b, VA.ShapeSheet.FormulaLiteral c, VA.ShapeSheet.FormulaLiteral d)
        {
            var row = new VA.ShapeGeometry.RowEllipticalArcTo(x, y, a, b, c, d);
            this.Rows.Add(row);
        }
    }
}