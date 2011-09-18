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
                row.AddTo(shape, update, row_count, sec_index);
                row_count++;
            }

            update.Execute(shape);
            return 0;
        }

        public void MoveTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y)
        {
            var row = VA.ShapeGeometry.GeometryRow.CreateMoveTo(x, y);
            this.Rows.Add(row);
        }

        public void LineTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y)
        {
            var row = VA.ShapeGeometry.GeometryRow.CreateLineTo(x, y);
            this.Rows.Add(row);
        }

        public void ArcTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a)
        {
            var row = VA.ShapeGeometry.GeometryRow.CreateArcTo(x, y, a);
            this.Rows.Add(row);
        }

        public void EllipticalArcTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a, VA.ShapeSheet.FormulaLiteral b, VA.ShapeSheet.FormulaLiteral c, VA.ShapeSheet.FormulaLiteral d)
        {
            var row = VA.ShapeGeometry.GeometryRow.CreateEllipticalArcTo(x, y, a,b,c,d);
            this.Rows.Add(row);
        }

        public void Ellipse(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a, VA.ShapeSheet.FormulaLiteral b, VA.ShapeSheet.FormulaLiteral c, VA.ShapeSheet.FormulaLiteral d)
        {
            var row = VA.ShapeGeometry.GeometryRow.CreateEllipse(x, y, a,b,c,d);
            this.Rows.Add(row);
        }

        public void NURBSTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a, VA.ShapeSheet.FormulaLiteral b, VA.ShapeSheet.FormulaLiteral c, VA.ShapeSheet.FormulaLiteral d, VA.ShapeSheet.FormulaLiteral e)
        {
            var row = VA.ShapeGeometry.GeometryRow.CreateNURBSTo(x, y, a,b,c,d,e);
            this.Rows.Add(row);
        }

    }
}