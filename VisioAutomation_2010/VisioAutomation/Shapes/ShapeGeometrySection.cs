using System.Collections;
using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Shapes
{
    public class ShapeGeometrySection : IEnumerable<ShapeGeometryRow>
    {
        private List<ShapeGeometryRow> Rows { get; }
        public Core.CellValue NoFill { get; set; }
        public Core.CellValue NoLine { get; set; }
        public Core.CellValue NoShow { get; set; }
        public Core.CellValue NoSnap { get; set; }
        public Core.CellValue NoQuickDrag { get; set; }

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

            var src_nofill = new VA.Core.Src(sec_index, 0, Core.SrcConstants.GeometryNoFill.Cell);
            var src_noline = new VA.Core.Src(sec_index, 0, Core.SrcConstants.GeometryNoLine.Cell);
            var src_noshow = new VA.Core.Src(sec_index, 0, Core.SrcConstants.GeometryNoShow.Cell);
            var src_nosnap = new VA.Core.Src(sec_index, 0, Core.SrcConstants.GeometryNoSnap.Cell);
            var src_noquickdrag = new VA.Core.Src(sec_index, 0, Core.SrcConstants.GeometryNoQuickDrag.Cell);

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

            writer.Commit(shape, Core.CellValueType.Formula);

            return 0;
        }

        public ShapeGeometryRow AddMoveTo(Core.CellValue x, Core.CellValue y)
        {
            var row = ShapeGeometryRow.CreateMoveTo(x, y);
            this.Rows.Add(row);
            return row;
        }

        public ShapeGeometryRow AddLineTo(Core.CellValue x, Core.CellValue y)
        {
            var row = ShapeGeometryRow.CreateLineTo(x, y);
            this.Rows.Add(row);
            return row;

        }

        public ShapeGeometryRow AddArcTo(Core.CellValue x, Core.CellValue y, Core.CellValue a)
        {
            var row = ShapeGeometryRow.CreateArcTo(x, y, a);
            this.Rows.Add(row);
            return row;

        }

        public ShapeGeometryRow AddEllipticalArcTo(Core.CellValue x, Core.CellValue y, Core.CellValue a, Core.CellValue b, Core.CellValue c, Core.CellValue d)
        {
            var row = ShapeGeometryRow.CreateEllipticalArcTo(x, y, a, b, c, d);
            this.Rows.Add(row);
            return row;

        }

        public ShapeGeometryRow AddEllipse(Core.CellValue x, Core.CellValue y, Core.CellValue a, Core.CellValue b, Core.CellValue c, Core.CellValue d)
        {
            var row = ShapeGeometryRow.CreateEllipse(x, y, a, b, c, d);
            this.Rows.Add(row);
            return row;

        }

        public ShapeGeometryRow AddNurbsTo(Core.CellValue x, Core.CellValue y, Core.CellValue a, Core.CellValue b, Core.CellValue c, Core.CellValue d, Core.CellValue e)
        {
            var row = ShapeGeometryRow.CreateNurbsTo(x, y, a, b, c, d, e);
            this.Rows.Add(row);
            return row;

        }

        public ShapeGeometryRow AddPolylineTo(Core.CellValue x, Core.CellValue y, Core.CellValue a)
        {
            var row = ShapeGeometryRow.CreatePolylineTo(x, y, a);
            this.Rows.Add(row);
            return row;

        }

        public ShapeGeometryRow AddInfiniteLine(Core.CellValue x, Core.CellValue y, Core.CellValue a, Core.CellValue b)
        {
            var row = ShapeGeometryRow.CreateInfiniteLine(x, y, a, b);
            this.Rows.Add(row);
            return row;

        }

        public ShapeGeometryRow AddSplineStart(Core.CellValue x, Core.CellValue y, Core.CellValue a, Core.CellValue b, Core.CellValue c, Core.CellValue d)
        {
            var row = ShapeGeometryRow.CreateSplineStart(x, y, a, b, c, d);
            this.Rows.Add(row);
            return row;

        }

        public ShapeGeometryRow AddSplineKnot(Core.CellValue x, Core.CellValue y, Core.CellValue a)
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