using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
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

            this.CanGlue = sec.AddCell(SRCCON.Controls_CanGlue, nameof(SRCCON.Controls_CanGlue));
            this.Tip = sec.AddCell(SRCCON.Controls_Tip, nameof(SRCCON.Controls_Tip));
            this.X = sec.AddCell(SRCCON.Controls_X, nameof(SRCCON.Controls_X));
            this.Y = sec.AddCell(SRCCON.Controls_Y, nameof(SRCCON.Controls_Y));
            this.YBehavior = sec.AddCell(SRCCON.Controls_YCon, nameof(SRCCON.Controls_YCon));
            this.XBehavior = sec.AddCell(SRCCON.Controls_XCon, nameof(SRCCON.Controls_XCon));
            this.XDynamics = sec.AddCell(SRCCON.Controls_XDyn, nameof(SRCCON.Controls_XDyn));
            this.YDynamics = sec.AddCell(SRCCON.Controls_YDyn, nameof(SRCCON.Controls_YDyn));

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