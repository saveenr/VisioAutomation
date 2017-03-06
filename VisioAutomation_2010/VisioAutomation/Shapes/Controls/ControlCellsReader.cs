using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.Controls
{
    class ControlCellsReader : MultiRowReader<Shapes.Controls.ControlCells>
    {
        public SubQueryColumn CanGlue { get; set; }
        public SubQueryColumn Tip { get; set; }
        public SubQueryColumn X { get; set; }
        public SubQueryColumn Y { get; set; }
        public SubQueryColumn YBehavior { get; set; }
        public SubQueryColumn XBehavior { get; set; }
        public SubQueryColumn XDynamics { get; set; }
        public SubQueryColumn YDynamics { get; set; }

        public ControlCellsReader()
        {
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionControls);

            this.CanGlue = sec.AddCell(SrcConstants.ControlCanGlue, nameof(SrcConstants.ControlCanGlue));
            this.Tip = sec.AddCell(SrcConstants.ControlTip, nameof(SrcConstants.ControlTip));
            this.X = sec.AddCell(SrcConstants.ControlX, nameof(SrcConstants.ControlX));
            this.Y = sec.AddCell(SrcConstants.ControlY, nameof(SrcConstants.ControlY));
            this.YBehavior = sec.AddCell(SrcConstants.ControlYCon, nameof(SrcConstants.ControlYCon));
            this.XBehavior = sec.AddCell(SrcConstants.ControlXCon, nameof(SrcConstants.ControlXCon));
            this.XDynamics = sec.AddCell(SrcConstants.ControlXDyn, nameof(SrcConstants.ControlXDyn));
            this.YDynamics = sec.AddCell(SrcConstants.ControlYDyn, nameof(SrcConstants.ControlYDyn));

        }

        public override Shapes.Controls.ControlCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Shapes.Controls.ControlCells();
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