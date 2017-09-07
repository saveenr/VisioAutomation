using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    class ControlCellsReader : ReaderMultiRow<ControlCells>
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
        {
            var sec = this.query.SectionQueries.Add(IVisio.VisSectionIndices.visSectionControls);

            this.CanGlue = sec.Columns.Add(SrcConstants.ControlCanGlue, nameof(SrcConstants.ControlCanGlue));
            this.Tip = sec.Columns.Add(SrcConstants.ControlTip, nameof(SrcConstants.ControlTip));
            this.X = sec.Columns.Add(SrcConstants.ControlX, nameof(SrcConstants.ControlX));
            this.Y = sec.Columns.Add(SrcConstants.ControlY, nameof(SrcConstants.ControlY));
            this.YBehavior = sec.Columns.Add(SrcConstants.ControlYBehavior, nameof(SrcConstants.ControlYBehavior));
            this.XBehavior = sec.Columns.Add(SrcConstants.ControlXBehavior, nameof(SrcConstants.ControlXBehavior));
            this.XDynamics = sec.Columns.Add(SrcConstants.ControlXDynamics, nameof(SrcConstants.ControlXDynamics));
            this.YDynamics = sec.Columns.Add(SrcConstants.ControlYDynamics, nameof(SrcConstants.ControlYDynamics));

        }

        public override ControlCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row)
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