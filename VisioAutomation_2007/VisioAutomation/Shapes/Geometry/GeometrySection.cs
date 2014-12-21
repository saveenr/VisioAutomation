using System.Collections;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Shapes.Geometry
{
    public class GeometrySection : IEnumerable<GeometryRow>
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

        public IEnumerator<GeometryRow> GetEnumerator()
        {
            foreach (var i in this.Rows)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()     
        {                                           
            return GetEnumerator();
        }

        public GeometryRow this[int index]
        {
            get { return this.Rows[index]; }
        }

        public short Render(IVisio.Shape shape)
        {
            short sec_index = GeometryHelper.AddSection(shape);
            short row_count = shape.RowCount[sec_index];

            var update = new VA.ShapeSheet.Update();

            var src_nofill = VA.ShapeSheet.SRCConstants.Geometry_NoFill.ForSectionAndRow(sec_index, 0);
            var src_noline = VA.ShapeSheet.SRCConstants.Geometry_NoLine.ForSectionAndRow(sec_index, 0);
            var src_noshow = VA.ShapeSheet.SRCConstants.Geometry_NoShow.ForSectionAndRow(sec_index, 0);
            var src_nosnap = VA.ShapeSheet.SRCConstants.Geometry_NoSnap.ForSectionAndRow(sec_index, 0);
            //var src_noquickdrag = VA.ShapeSheet.SRCConstants.Geometry_NoQuickDrag.ForSectionAndRow(sec_index, 0);

            update.SetFormulaIgnoreNull(src_nofill, this.NoFill);
            update.SetFormulaIgnoreNull(src_noline, this.NoLine);
            update.SetFormulaIgnoreNull(src_noshow, this.NoShow);
            update.SetFormulaIgnoreNull(src_nosnap, this.NoSnap);
            //update.SetFormulaIgnoreNull(src_noquickdrag, this.NoQuickDrag);

            foreach (var row in this.Rows)
            {
                row.AddTo(shape, update, row_count, sec_index);
                row_count++;
            }

            update.Execute(shape);
            return 0;
        }

        public GeometryRow AddMoveTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y)
        {
            var row = GeometryRow.CreateMoveTo(x, y);
            this.Rows.Add(row);
            return row;
        }

        public GeometryRow AddLineTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y)
        {
            var row = GeometryRow.CreateLineTo(x, y);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddArcTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a)
        {
            var row = GeometryRow.CreateArcTo(x, y, a);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddEllipticalArcTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a, VA.ShapeSheet.FormulaLiteral b, VA.ShapeSheet.FormulaLiteral c, VA.ShapeSheet.FormulaLiteral d)
        {
            var row = GeometryRow.CreateEllipticalArcTo(x, y, a, b, c, d);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddEllipse(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a, VA.ShapeSheet.FormulaLiteral b, VA.ShapeSheet.FormulaLiteral c, VA.ShapeSheet.FormulaLiteral d)
        {
            var row = GeometryRow.CreateEllipse(x, y, a, b, c, d);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddNURBSTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a, VA.ShapeSheet.FormulaLiteral b, VA.ShapeSheet.FormulaLiteral c, VA.ShapeSheet.FormulaLiteral d, VA.ShapeSheet.FormulaLiteral e)
        {
            var row = GeometryRow.CreateNURBSTo(x, y, a, b, c, d, e);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddPolylineTo(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a)
        {
            var row = GeometryRow.CreatePolylineTo(x, y, a);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddInfiniteLine(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a, VA.ShapeSheet.FormulaLiteral b)
        {
            var row = GeometryRow.CreateInfiniteLine(x, y, a, b);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddSplineStart(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a, VA.ShapeSheet.FormulaLiteral b, VA.ShapeSheet.FormulaLiteral c, VA.ShapeSheet.FormulaLiteral d)
        {
            var row = GeometryRow.CreateSplineStart(x, y, a, b, c, d);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddSplineKnot(VA.ShapeSheet.FormulaLiteral x, VA.ShapeSheet.FormulaLiteral y, VA.ShapeSheet.FormulaLiteral a)
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