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
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionControls);

            this.CanGlue = sec.AddColumn(SrcConstants.ControlCanGlue, nameof(SrcConstants.ControlCanGlue));
            this.Tip = sec.AddColumn(SrcConstants.ControlTip, nameof(SrcConstants.ControlTip));
            this.X = sec.AddColumn(SrcConstants.ControlX, nameof(SrcConstants.ControlX));
            this.Y = sec.AddColumn(SrcConstants.ControlY, nameof(SrcConstants.ControlY));
            this.YBehavior = sec.AddColumn(SrcConstants.ControlYBehavior, nameof(SrcConstants.ControlYBehavior));
            this.XBehavior = sec.AddColumn(SrcConstants.ControlXBehavior, nameof(SrcConstants.ControlXBehavior));
            this.XDynamics = sec.AddColumn(SrcConstants.ControlXDynamics, nameof(SrcConstants.ControlXDynamics));
            this.YDynamics = sec.AddColumn(SrcConstants.ControlYDynamics, nameof(SrcConstants.ControlYDynamics));

        }

        public override ControlCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
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