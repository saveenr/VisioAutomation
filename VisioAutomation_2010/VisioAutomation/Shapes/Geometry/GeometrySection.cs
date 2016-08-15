using System.Collections;
using System.Collections.Generic;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.Geometry
{
    public class GeometrySection : IEnumerable<GeometryRow>
    {
        private List<GeometryRow> Rows { get; }
        public ShapeSheet.FormulaLiteral NoFill { get; set; }
        public ShapeSheet.FormulaLiteral NoLine { get; set; }
        public ShapeSheet.FormulaLiteral NoShow { get; set; }
        public ShapeSheet.FormulaLiteral NoSnap { get; set; }
        public ShapeSheet.FormulaLiteral NoQuickDrag { get; set; }

        public GeometrySection()
        {
            this.Rows = new List<GeometryRow>();
        }

        public IEnumerator<GeometryRow> GetEnumerator()
        {
            foreach (var i in this.Rows)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()     
        {                                           
            return this.GetEnumerator();
        }

        public GeometryRow this[int index]
        {
            get { return this.Rows[index]; }
        }

        public short Render(IVisio.Shape shape)
        {
            short sec_index = GeometryHelper.AddSection(shape);
            short row_count = shape.RowCount[sec_index];

            var update = new FormulaWriterSRC();

            var src_nofill = ShapeSheet.SRCConstants.Geometry_NoFill.ForSectionAndRow(sec_index, 0);
            var src_noline = ShapeSheet.SRCConstants.Geometry_NoLine.ForSectionAndRow(sec_index, 0);
            var src_noshow = ShapeSheet.SRCConstants.Geometry_NoShow.ForSectionAndRow(sec_index, 0);
            var src_nosnap = ShapeSheet.SRCConstants.Geometry_NoSnap.ForSectionAndRow(sec_index, 0);
            var src_noquickdrag = ShapeSheet.SRCConstants.Geometry_NoQuickDrag.ForSectionAndRow(sec_index, 0);

            update.SetFormula(src_nofill, this.NoFill);
            update.SetFormula(src_noline, this.NoLine);
            update.SetFormula(src_noshow, this.NoShow);
            update.SetFormula(src_nosnap, this.NoSnap);
            update.SetFormula(src_noquickdrag, this.NoQuickDrag);

            foreach (var row in this.Rows)
            {
                row.AddTo(shape, update, row_count, sec_index);
                row_count++;
            }

            update.Commit(shape);
            return 0;
        }

        public GeometryRow AddMoveTo(ShapeSheet.FormulaLiteral x, ShapeSheet.FormulaLiteral y)
        {
            var row = GeometryRow.CreateMoveTo(x, y);
            this.Rows.Add(row);
            return row;
        }

        public GeometryRow AddLineTo(ShapeSheet.FormulaLiteral x, ShapeSheet.FormulaLiteral y)
        {
            var row = GeometryRow.CreateLineTo(x, y);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddArcTo(ShapeSheet.FormulaLiteral x, ShapeSheet.FormulaLiteral y, ShapeSheet.FormulaLiteral a)
        {
            var row = GeometryRow.CreateArcTo(x, y, a);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddEllipticalArcTo(ShapeSheet.FormulaLiteral x, ShapeSheet.FormulaLiteral y, ShapeSheet.FormulaLiteral a, ShapeSheet.FormulaLiteral b, ShapeSheet.FormulaLiteral c, ShapeSheet.FormulaLiteral d)
        {
            var row = GeometryRow.CreateEllipticalArcTo(x, y, a, b, c, d);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddEllipse(ShapeSheet.FormulaLiteral x, ShapeSheet.FormulaLiteral y, ShapeSheet.FormulaLiteral a, ShapeSheet.FormulaLiteral b, ShapeSheet.FormulaLiteral c, ShapeSheet.FormulaLiteral d)
        {
            var row = GeometryRow.CreateEllipse(x, y, a, b, c, d);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddNURBSTo(ShapeSheet.FormulaLiteral x, ShapeSheet.FormulaLiteral y, ShapeSheet.FormulaLiteral a, ShapeSheet.FormulaLiteral b, ShapeSheet.FormulaLiteral c, ShapeSheet.FormulaLiteral d, ShapeSheet.FormulaLiteral e)
        {
            var row = GeometryRow.CreateNURBSTo(x, y, a, b, c, d, e);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddPolylineTo(ShapeSheet.FormulaLiteral x, ShapeSheet.FormulaLiteral y, ShapeSheet.FormulaLiteral a)
        {
            var row = GeometryRow.CreatePolylineTo(x, y, a);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddInfiniteLine(ShapeSheet.FormulaLiteral x, ShapeSheet.FormulaLiteral y, ShapeSheet.FormulaLiteral a, ShapeSheet.FormulaLiteral b)
        {
            var row = GeometryRow.CreateInfiniteLine(x, y, a, b);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddSplineStart(ShapeSheet.FormulaLiteral x, ShapeSheet.FormulaLiteral y, ShapeSheet.FormulaLiteral a, ShapeSheet.FormulaLiteral b, ShapeSheet.FormulaLiteral c, ShapeSheet.FormulaLiteral d)
        {
            var row = GeometryRow.CreateSplineStart(x, y, a, b, c, d);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddSplineKnot(ShapeSheet.FormulaLiteral x, ShapeSheet.FormulaLiteral y, ShapeSheet.FormulaLiteral a)
        {
            var row = GeometryRow.CreateSplineKnot(x, y, a);
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