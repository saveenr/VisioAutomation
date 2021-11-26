using System.Collections;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class GeometrySection : IEnumerable<GeometryRow>
    {
        private List<GeometryRow> Rows { get; }
        public Core.CellValue NoFill { get; set; }
        public Core.CellValue NoLine { get; set; }
        public Core.CellValue NoShow { get; set; }
        public Core.CellValue NoSnap { get; set; }
        public Core.CellValue NoQuickDrag { get; set; }

        public GeometrySection()
        {
            this.Rows = new List<GeometryRow>();
        }

        public IEnumerator<GeometryRow> GetEnumerator()
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

        public GeometryRow this[int index]
        {
            get { return this.Rows[index]; }
        }

        public short Render(IVisio.Shape shape)
        {
            short sec_index = GeometryHelper.AddSection(shape);
            short row_count = shape.RowCount[sec_index];

            var writer = new ShapeSheet.Writers.SrcWriter();

            var src_nofill = new Core.Src(sec_index, 0, Core.SrcConstants.GeometryNoFill.Cell);
            var src_noline = new Core.Src(sec_index, 0, Core.SrcConstants.GeometryNoLine.Cell);
            var src_noshow = new Core.Src(sec_index, 0, Core.SrcConstants.GeometryNoShow.Cell);
            var src_nosnap = new Core.Src(sec_index, 0, Core.SrcConstants.GeometryNoSnap.Cell);
            var src_noquickdrag = new Core.Src(sec_index, 0, Core.SrcConstants.GeometryNoQuickDrag.Cell);

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

        public GeometryRow AddMoveTo(Core.CellValue x, Core.CellValue y)
        {
            var row = GeometryRow.CreateMoveTo(x, y);
            this.Rows.Add(row);
            return row;
        }

        public GeometryRow AddLineTo(Core.CellValue x, Core.CellValue y)
        {
            var row = GeometryRow.CreateLineTo(x, y);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddArcTo(Core.CellValue x, Core.CellValue y, Core.CellValue a)
        {
            var row = GeometryRow.CreateArcTo(x, y, a);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddEllipticalArcTo(Core.CellValue x, Core.CellValue y, Core.CellValue a, Core.CellValue b, Core.CellValue c, Core.CellValue d)
        {
            var row = GeometryRow.CreateEllipticalArcTo(x, y, a, b, c, d);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddEllipse(Core.CellValue x, Core.CellValue y, Core.CellValue a, Core.CellValue b, Core.CellValue c, Core.CellValue d)
        {
            var row = GeometryRow.CreateEllipse(x, y, a, b, c, d);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddNurbsTo(Core.CellValue x, Core.CellValue y, Core.CellValue a, Core.CellValue b, Core.CellValue c, Core.CellValue d, Core.CellValue e)
        {
            var row = GeometryRow.CreateNurbsTo(x, y, a, b, c, d, e);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddPolylineTo(Core.CellValue x, Core.CellValue y, Core.CellValue a)
        {
            var row = GeometryRow.CreatePolylineTo(x, y, a);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddInfiniteLine(Core.CellValue x, Core.CellValue y, Core.CellValue a, Core.CellValue b)
        {
            var row = GeometryRow.CreateInfiniteLine(x, y, a, b);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddSplineStart(Core.CellValue x, Core.CellValue y, Core.CellValue a, Core.CellValue b, Core.CellValue c, Core.CellValue d)
        {
            var row = GeometryRow.CreateSplineStart(x, y, a, b, c, d);
            this.Rows.Add(row);
            return row;

        }

        public GeometryRow AddSplineKnot(Core.CellValue x, Core.CellValue y, Core.CellValue a)
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