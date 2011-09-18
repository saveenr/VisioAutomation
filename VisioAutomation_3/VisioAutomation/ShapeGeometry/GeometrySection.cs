using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeGeometry
{
    public class GeometrySection
    {
        private List<VA.ShapeGeometry.GeometryRow> Rows;

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

            var src_nofill = new VA.ShapeSheet.SRC(sec_index, 0, VA.ShapeSheet.SRCConstants.Geometry_NoFill.Cell);
            var src_noline = new VA.ShapeSheet.SRC(sec_index, 0, VA.ShapeSheet.SRCConstants.Geometry_NoLine.Cell);
            var src_noshow = new VA.ShapeSheet.SRC(sec_index, 0, VA.ShapeSheet.SRCConstants.Geometry_NoShow.Cell);
            var src_nosnap = new VA.ShapeSheet.SRC(sec_index, 0, VA.ShapeSheet.SRCConstants.Geometry_NoSnap.Cell);

            update.SetFormulaIgnoreNull(src_nofill, this.NoFill);
            update.SetFormulaIgnoreNull(src_noline, this.NoLine);
            update.SetFormulaIgnoreNull(src_noshow, this.NoShow);
            update.SetFormulaIgnoreNull(src_nosnap, this.NoSnap);

            foreach (var row in this.Rows)
            {
                if (row is RowMoveTo)
                {
                    var moveto_row = (RowMoveTo) row;
                    short row_index = shape.AddRow(sec_index, row_count, (short)IVisio.VisRowTags.visTagMoveTo);
                    var x_src = new VA.ShapeSheet.SRC(sec_index, row_index, VA.ShapeSheet.SRCConstants.Geometry_X.Cell);
                    var y_src = new VA.ShapeSheet.SRC(sec_index, row_index, VA.ShapeSheet.SRCConstants.Geometry_Y.Cell);
                    update.SetFormulaIgnoreNull(x_src, moveto_row.X);
                    update.SetFormulaIgnoreNull(y_src, moveto_row.Y);
                    row_count++;
                }
                else if (row is RowLineTo)
                {
                    var lineto_row = (RowLineTo)row;
                    short row_index = shape.AddRow(sec_index, row_count, (short)IVisio.VisRowTags.visTagLineTo);
                    var x_src = new VA.ShapeSheet.SRC(sec_index, row_index, VA.ShapeSheet.SRCConstants.Geometry_X.Cell);
                    var y_src = new VA.ShapeSheet.SRC(sec_index, row_index, VA.ShapeSheet.SRCConstants.Geometry_Y.Cell);
                    update.SetFormulaIgnoreNull(x_src, lineto_row.X);
                    update.SetFormulaIgnoreNull(y_src, lineto_row.Y);
                    row_count++;
                }
                else if (row is RowArcTo)
                {
                    var arcto_row = (RowArcTo)row;
                    short row_index = shape.AddRow(sec_index, row_count, (short)IVisio.VisRowTags.visTagArcTo);
                    var x_src = new VA.ShapeSheet.SRC(sec_index, row_index, VA.ShapeSheet.SRCConstants.Geometry_X.Cell);
                    var y_src = new VA.ShapeSheet.SRC(sec_index, row_index, VA.ShapeSheet.SRCConstants.Geometry_Y.Cell);
                    var a_src = new VA.ShapeSheet.SRC(sec_index, row_index, VA.ShapeSheet.SRCConstants.Geometry_A.Cell);
                    update.SetFormulaIgnoreNull(x_src, arcto_row.X);
                    update.SetFormulaIgnoreNull(y_src, arcto_row.Y);
                    update.SetFormulaIgnoreNull(a_src, arcto_row.A);
                    row_count++;
                }
                else if (row is RowEllipticalArcTo)
                {
                    var arcto_row = (RowEllipticalArcTo)row;
                    short row_index = shape.AddRow(sec_index, row_count, (short)IVisio.VisRowTags.visTagArcTo);
                    var x_src = new VA.ShapeSheet.SRC(sec_index, row_index, VA.ShapeSheet.SRCConstants.Geometry_X.Cell);
                    var y_src = new VA.ShapeSheet.SRC(sec_index, row_index, VA.ShapeSheet.SRCConstants.Geometry_Y.Cell);
                    var a_src = new VA.ShapeSheet.SRC(sec_index, row_index, VA.ShapeSheet.SRCConstants.Geometry_A.Cell);
                    var b_src = new VA.ShapeSheet.SRC(sec_index, row_index, VA.ShapeSheet.SRCConstants.Geometry_B.Cell);
                    var c_src = new VA.ShapeSheet.SRC(sec_index, row_index, VA.ShapeSheet.SRCConstants.Geometry_C.Cell);
                    var d_src = new VA.ShapeSheet.SRC(sec_index, row_index, VA.ShapeSheet.SRCConstants.Geometry_D.Cell);
                    update.SetFormulaIgnoreNull(x_src, arcto_row.X);
                    update.SetFormulaIgnoreNull(y_src, arcto_row.Y);
                    update.SetFormulaIgnoreNull(a_src, arcto_row.A);
                    update.SetFormulaIgnoreNull(b_src, arcto_row.B);
                    update.SetFormulaIgnoreNull(c_src, arcto_row.C);
                    update.SetFormulaIgnoreNull(d_src, arcto_row.D);
                    row_count++;
                }
                else
                {
                    string msg = string.Format("Unsupported row type \"{0}\"", row.GetType().Name);
                    throw new AutomationException(msg);
                }
            }

            update.Execute(shape);
            return 0;
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