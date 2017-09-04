using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    class ControlCellsReader : ReaderMultiRow<ControlCells>
    {
        public SectionSubQueryColumn CanGlue { get; set; }
        public SectionSubQueryColumn Tip { get; set; }
        public SectionSubQueryColumn X { get; set; }
        public SectionSubQueryColumn Y { get; set; }
        public SectionSubQueryColumn YBehavior { get; set; }
        public SectionSubQueryColumn XBehavior { get; set; }
        public SectionSubQueryColumn XDynamics { get; set; }
        public SectionSubQueryColumn YDynamics { get; set; }

        public ControlCellsReader()
        {
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionControls);

            this.CanGlue = sec.AddCell(SrcConstants.ControlCanGlue, nameof(SrcConstants.ControlCanGlue));
            this.Tip = sec.AddCell(SrcConstants.ControlTip, nameof(SrcConstants.ControlTip));
            this.X = sec.AddCell(SrcConstants.ControlX, nameof(SrcConstants.ControlX));
            this.Y = sec.AddCell(SrcConstants.ControlY, nameof(SrcConstants.ControlY));
            this.YBehavior = sec.AddCell(SrcConstants.ControlYBehavior, nameof(SrcConstants.ControlYBehavior));
            this.XBehavior = sec.AddCell(SrcConstants.ControlXBehavior, nameof(SrcConstants.ControlXBehavior));
            this.XDynamics = sec.AddCell(SrcConstants.ControlXDynamics, nameof(SrcConstants.ControlXDynamics));
            this.YDynamics = sec.AddCell(SrcConstants.ControlYDynamics, nameof(SrcConstants.ControlYDynamics));

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