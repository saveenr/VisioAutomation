using System.Collections;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Shapes
{
    public class ShapeGeometrySection : IEnumerable<ShapeGeometryRow>
    {
        private List<ShapeGeometryRow> Rows { get; }
        public ShapeSheet.CellValueLiteral NoFill { get; set; }
        public ShapeSheet.CellValueLiteral NoLine { get; set; }
        public ShapeSheet.CellValueLiteral NoShow { get; set; }
        public ShapeSheet.CellValueLiteral NoSnap { get; set; }
        public ShapeSheet.CellValueLiteral NoQuickDrag { get; set; }

        public ShapeGeometrySection()
        {
            this.Rows = new List<ShapeGeometryRow>();
        }

        public IEnumerator<ShapeGeometryRow> GetEnumerator()
        {
            foreach (var row in this.Rows)
            {
                yield return row;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()     
        {                                           
            return this.GetEnumerator();
        }

        public ShapeGeometryRow this[int index]
        {
            get { return this.Rows[index]; }
        }

        public short Render(IVisio.Shape shape)
        {
            short sec_index = ShapeGeometryHelper.AddSection(shape);
            short row_count = shape.RowCount[sec_index];

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();

            var src_nofill = new VA.ShapeSheet.Src(sec_index, 0, ShapeSheet.SrcConstants.GeometryNoFill.Cell);
            var src_noline = new VA.ShapeSheet.Src(sec_index, 0, ShapeSheet.SrcConstants.GeometryNoLine.Cell);
            var src_noshow = new VA.ShapeSheet.Src(sec_index, 0, ShapeSheet.SrcConstants.GeometryNoShow.Cell);
            var src_nosnap = new VA.ShapeSheet.Src(sec_index, 0, ShapeSheet.SrcConstants.GeometryNoSnap.Cell);
            var src_noquickdrag = new VA.ShapeSheet.Src(sec_index, 0, ShapeSheet.SrcConstants.GeometryNoQuickDrag.Cell);

            writer.SetValue(src_nofill, this.NoFill);
            writer.SetValue(src_noline, this.NoLine);
            writer.SetValue(src_noshow, this.NoShow);
            writer.SetValue(src_nosnap, this.NoSnap);
            writer.SetValue(src_noquickdrag, this.NoQuickDrag);

            foreach (var row in this.Rows)
            {
                row.AddTo(shape, writer, row_count, sec_index);
                row_count++;
            }

            writer.CommitFormulas(shape);

            return 0;
        }

        public ShapeGeometryRow AddMoveTo(ShapeSheet.CellValueLiteral x, ShapeSheet.CellValueLiteral y)
        {
            var row = ShapeGeometryRow.CreateMoveTo(x, y);
            this.Rows.Add(row);
            return row;
        }

        public ShapeGeometryRow AddLineTo(ShapeSheet.CellValueLiteral x, ShapeSheet.CellValueLiteral y)
        {
            var row = ShapeGeometryRow.CreateLineTo(x, y);
            this.Rows.Add(row);
            return row;

        }

        public ShapeGeometryRow AddArcTo(ShapeSheet.CellValueLiteral x, ShapeSheet.CellValueLiteral y, ShapeSheet.CellValueLiteral a)
        {
            var row = ShapeGeometryRow.CreateArcTo(x, y, a);
            this.Rows.Add(row);
            return row;

        }

        public ShapeGeometryRow AddEllipticalArcTo(ShapeSheet.CellValueLiteral x, ShapeSheet.CellValueLiteral y, ShapeSheet.CellValueLiteral a, ShapeSheet.CellValueLiteral b, ShapeSheet.CellValueLiteral c, ShapeSheet.CellValueLiteral d)
        {
            var row = ShapeGeometryRow.CreateEllipticalArcTo(x, y, a, b, c, d);
            this.Rows.Add(row);
            return row;

        }

        public ShapeGeometryRow AddEllipse(ShapeSheet.CellValueLiteral x, ShapeSheet.CellValueLiteral y, ShapeSheet.CellValueLiteral a, ShapeSheet.CellValueLiteral b, ShapeSheet.CellValueLiteral c, ShapeSheet.CellValueLiteral d)
        {
            var row = ShapeGeometryRow.CreateEllipse(x, y, a, b, c, d);
            this.Rows.Add(row);
            return row;

        }

        public ShapeGeometryRow AddNURBSTo(ShapeSheet.CellValueLiteral x, ShapeSheet.CellValueLiteral y, ShapeSheet.CellValueLiteral a, ShapeSheet.CellValueLiteral b, ShapeSheet.CellValueLiteral c, ShapeSheet.CellValueLiteral d, ShapeSheet.CellValueLiteral e)
        {
            var row = ShapeGeometryRow.CreateNURBSTo(x, y, a, b, c, d, e);
            this.Rows.Add(row);
            return row;

        }

        public ShapeGeometryRow AddPolylineTo(ShapeSheet.CellValueLiteral x, ShapeSheet.CellValueLiteral y, ShapeSheet.CellValueLiteral a)
        {
            var row = ShapeGeometryRow.CreatePolylineTo(x, y, a);
            this.Rows.Add(row);
            return row;

        }

        public ShapeGeometryRow AddInfiniteLine(ShapeSheet.CellValueLiteral x, ShapeSheet.CellValueLiteral y, ShapeSheet.CellValueLiteral a, ShapeSheet.CellValueLiteral b)
        {
            var row = ShapeGeometryRow.CreateInfiniteLine(x, y, a, b);
            this.Rows.Add(row);
            return row;

        }

        public ShapeGeometryRow AddSplineStart(ShapeSheet.CellValueLiteral x, ShapeSheet.CellValueLiteral y, ShapeSheet.CellValueLiteral a, ShapeSheet.CellValueLiteral b, ShapeSheet.CellValueLiteral c, ShapeSheet.CellValueLiteral d)
        {
            var row = ShapeGeometryRow.CreateSplineStart(x, y, a, b, c, d);
            this.Rows.Add(row);
            return row;

        }

        public ShapeGeometryRow AddSplineKnot(ShapeSheet.CellValueLiteral x, ShapeSheet.CellValueLiteral y, ShapeSheet.CellValueLiteral a)
        {
            var row = ShapeGeometryRow.CreateSplineKnot(x, y, a);
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