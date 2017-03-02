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

            this.CanGlue = sec.AddCell(SrcConstants.Controls_CanGlue, nameof(SrcConstants.Controls_CanGlue));
            this.Tip = sec.AddCell(SrcConstants.Controls_Tip, nameof(SrcConstants.Controls_Tip));
            this.X = sec.AddCell(SrcConstants.Controls_X, nameof(SrcConstants.Controls_X));
            this.Y = sec.AddCell(SrcConstants.Controls_Y, nameof(SrcConstants.Controls_Y));
            this.YBehavior = sec.AddCell(SrcConstants.Controls_YCon, nameof(SrcConstants.Controls_YCon));
            this.XBehavior = sec.AddCell(SrcConstants.Controls_XCon, nameof(SrcConstants.Controls_XCon));
            this.XDynamics = sec.AddCell(SrcConstants.Controls_XDyn, nameof(SrcConstants.Controls_XDyn));
            this.YDynamics = sec.AddCell(SrcConstants.Controls_YDyn, nameof(SrcConstants.Controls_YDyn));

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