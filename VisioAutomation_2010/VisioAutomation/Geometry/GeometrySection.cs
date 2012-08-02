using System.Collections;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Geometry
{
    public class GeometrySection : IEnumerable<VA.Geometry.GeometryRow>
    {
        private List<GeometryRow> Rows { get; set; }
        public VA.ShapeSheet.FormulaLiteral NoFill { get; set; }
        public VA.ShapeSheet.FormulaLiteral NoLine { get; set; }
        public VA.ShapeSheet.FormulaLiteral NoShow { get; set; }
        public VA.ShapeSheet.FormulaLiteral NoSnap { get; set; }
        public VA.ShapeSheet.FormulaLiteral NoQuickDrag { get; set; }

        public GeometrySection()
        {
            this.Rows = new List<GeometryRow>();
        }

        public IEnumerator<VA.Geometry.GeometryRow> GetEnumerator()
        {
            foreach (var i in this.Rows)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()     // Explicit implementation
        {                                           // keeps it hidden.
            return GetEnumerator();
        }

        public VA.Geometry.GeometryRow this[int index]
        {
            get { return this.Rows[index]; }
        }

        public short Render(IVisio.Shape shape)
        {
            short sec_index = GeometryHelper.AddGeometrySection(shape);
            short row_count = shape.RowCount[sec_index];

            var update = new VA.ShapeSheet.Update.UpdateBase();

            var src_nofill = VA.ShapeSheet.SRCConstants.Geometry_NoFill.ForSectionAndRow(sec_index, 0);
            var src_noline = VA.ShapeSheet.SRCConstants.Geometry_NoLine.ForSectionAndRow(sec_index, 0);
            var src_noshow = VA.ShapeSheet.SRCConstants.Geometry_NoShow.ForSectionAndRow(sec_index, 0);
            var src_nosnap = VA.ShapeSheet.SRCConstants.Geometry_NoSnap.ForSectionAndRow(sec_index, 0);
            var src_noquickdrag = VA.ShapeSheet.SRCConstants.Geometry_NoQuickDrag.ForSectionAndRow(sec_index, 0);

            update.SetFormulaIgnoreNull(src_nofill, this.NoFill);
            update.SetFormulaIgnoreNull(src_noline, this.NoLine);
            update.SetFormulaIgnoreNull(src_noshow, this.NoShow);
            update.SetFormulaIgnoreNull(src_nosnap, this.NoSnap);
            update.SetFormulaIgnoreNull(src_noquickdrag, this.NoQuickDrag);

            foreach (var row in this.Rows)
            {
                row.AddTo(shape, update, row_count, sec_index);
                row_count++;
            }

            update.Execute(shape);
            return 0;
        }

        public VA.Geometry.GeometryRow AddMoveTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y)
        {
            var row = VA.Geometry.GeometryRow.CreateMoveTo(x, y);
            this.Rows.Add(row);
            return row;
        }

        public VA.Geometry.GeometryRow AddLineTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y)
        {
            var row = VA.Geometry.GeometryRow.CreateLineTo(x, y);
            this.Rows.Add(row);
            return row;

        }

        public VA.Geometry.GeometryRow AddArcTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a)
        {
            var row = VA.Geometry.GeometryRow.CreateArcTo(x, y, a);
            this.Rows.Add(row);
            return row;

        }

        public VA.Geometry.GeometryRow AddEllipticalArcTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a, VA.ShapeSheet.FormulaLiteral b, VA.ShapeSheet.FormulaLiteral c, VA.ShapeSheet.FormulaLiteral d)
        {
            var row = VA.Geometry.GeometryRow.CreateEllipticalArcTo(x, y, a, b, c, d);
            this.Rows.Add(row);
            return row;

        }

        public VA.Geometry.GeometryRow AddEllipse(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a, VA.ShapeSheet.FormulaLiteral b, VA.ShapeSheet.FormulaLiteral c, VA.ShapeSheet.FormulaLiteral d)
        {
            var row = VA.Geometry.GeometryRow.CreateEllipse(x, y, a, b, c, d);
            this.Rows.Add(row);
            return row;

        }

        public VA.Geometry.GeometryRow AddNURBSTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a, VA.ShapeSheet.FormulaLiteral b, VA.ShapeSheet.FormulaLiteral c, VA.ShapeSheet.FormulaLiteral d, VA.ShapeSheet.FormulaLiteral e)
        {
            var row = VA.Geometry.GeometryRow.CreateNURBSTo(x, y, a, b, c, d, e);
            this.Rows.Add(row);
            return row;

        }

        public VA.Geometry.GeometryRow AddPolylineTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a)
        {
            var row = VA.Geometry.GeometryRow.CreatePolylineTo(x, y, a);
            this.Rows.Add(row);
            return row;

        }

        public VA.Geometry.GeometryRow AddInfiniteLine(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a, VA.ShapeSheet.FormulaLiteral b)
        {
            var row = VA.Geometry.GeometryRow.CreateInfiniteLine(x, y, a, b);
            this.Rows.Add(row);
            return row;

        }

        public VA.Geometry.GeometryRow AddSplineStart(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a, VA.ShapeSheet.FormulaLiteral b, VA.ShapeSheet.FormulaLiteral c, VA.ShapeSheet.FormulaLiteral d)
        {
            var row = VA.Geometry.GeometryRow.CreateSplineStart(x, y, a, b, c, d);
            this.Rows.Add(row);
            return row;

        }

        public VA.Geometry.GeometryRow AddSplineKnot(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a)
        {
            var row = VA.Geometry.GeometryRow.CreateSplineKnot(x, y, a);
            this.Rows.Add(row);
            return row;
        }

        public int Count
        {
            get { return this.Rows.Count; }
        }

        public void Clear()
        {
            this.Rows.Clear();
        }
    }
}