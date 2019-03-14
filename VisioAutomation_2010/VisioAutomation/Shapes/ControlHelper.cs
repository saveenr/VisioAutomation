using System;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public static class ControlHelper
    {
        public static int Add(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            var ctrl = new ControlCells();

            return ControlHelper.Add(shape, ctrl);
        }

        public static int Add(IVisio.Shape shape, ControlCells ctrl)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            short row = shape.AddRow((short)IVisio.VisSectionIndices.visSectionControls,
                                     (short)IVisio.VisRowIndices.visRowLast,
                                     (short)IVisio.VisRowTags.visTagDefault);

            ControlHelper.Set(shape, row, ctrl);

            return row;
        }

        public static int Set( IVisio.Shape shape, short row, ControlCells ctrl)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }


            if (ctrl.XDynamics.Value==null)
            {
                ctrl.XDynamics = string.Format("Controls.Row_{0}", row + 1);
            }

            if (ctrl.YDynamics.Value == null)
            {
                ctrl.YDynamics = string.Format("Controls.Row_{0}.Y", row + 1);
            }

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetFormulas(ctrl, row);

            writer.Commit(shape);

            return row;
        }

        public static void Delete(IVisio.Shape shape, int index)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            if (index < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            var row = (IVisio.VisRowIndices)index;
            shape.DeleteRow( (short) IVisio.VisSectionIndices.visSectionControls, (short)row);
        }

        public static int GetCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            return shape.RowCount[(short)IVisio.VisSectionIndices.visSectionControls];
        }

        public static List<List<ControlCells>> GetControlCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = ControlCells_lazy_reader.Value;
            return reader.GetCellsMultiRow(page, shapeids, type);
        }

        public static List<ControlCells> GetControlCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = ControlCells_lazy_reader.Value;
            return reader.GetCellsMultiRow(shape, type);
        }

        private static readonly System.Lazy<ControlCellsReader> ControlCells_lazy_reader = new System.Lazy<ControlCellsReader>();

        class ControlCellsReader : CellGroupReader<ControlCells>
        {
            public SectionQueryColumn CanGlue { get; set; }
            public SectionQueryColumn Tip { get; set; }
            public SectionQueryColumn X { get; set; }
            public SectionQueryColumn Y { get; set; }
            public SectionQueryColumn YBehavior { get; set; }
            public SectionQueryColumn XBehavior { get; set; }
            public SectionQueryColumn XDynamics { get; set; }
            public SectionQueryColumn YDynamics { get; set; }

            public ControlCellsReader()
                : base(new VisioAutomation.ShapeSheet.Query.SectionsQuery())

            {
                var sec = this.query_multirow.SectionQueries.Add(IVisio.VisSectionIndices.visSectionControls);

                this.CanGlue = sec.Columns.Add(SrcConstants.ControlCanGlue, nameof(this.CanGlue));
                this.Tip = sec.Columns.Add(SrcConstants.ControlTip, nameof(this.Tip));
                this.X = sec.Columns.Add(SrcConstants.ControlX, nameof(this.X));
                this.Y = sec.Columns.Add(SrcConstants.ControlY, nameof(this.Y));
                this.YBehavior = sec.Columns.Add(SrcConstants.ControlYBehavior, nameof(this.YBehavior));
                this.XBehavior = sec.Columns.Add(SrcConstants.ControlXBehavior, nameof(this.XBehavior));
                this.XDynamics = sec.Columns.Add(SrcConstants.ControlXDynamics, nameof(this.XDynamics));
                this.YDynamics = sec.Columns.Add(SrcConstants.ControlYDynamics, nameof(this.YDynamics));

            }

            public override ControlCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new ControlCells();
                cells.CanGlue = row[this.CanGlue];
                cells.Tip = row[this.Tip];
                cells.X = row[this.X];
                cells.Y = row[this.Y];
                cells.YBehavior = row[this.YBehavior];
                cells.XBehavior = row[this.XBehavior];
                cells.XDynamics = row[this.XDynamics];
                cells.YDynamics = row[this.YDynamics];
                return cells;
            }
        }

    }
}