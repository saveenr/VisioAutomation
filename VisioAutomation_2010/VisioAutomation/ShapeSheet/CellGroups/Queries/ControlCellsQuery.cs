using VisioAutomation.ShapeSheet.Queries.Columns;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.ShapeSheet.CellGroups.Queries
{
    class ControlCellsQuery : CellGroupMultiRowQuery<Shapes.Controls.ControlCells>
    {
        public ColumnSubQuery CanGlue { get; set; }
        public ColumnSubQuery Tip { get; set; }
        public ColumnSubQuery X { get; set; }
        public ColumnSubQuery Y { get; set; }
        public ColumnSubQuery YBehavior { get; set; }
        public ColumnSubQuery XBehavior { get; set; }
        public ColumnSubQuery XDynamics { get; set; }
        public ColumnSubQuery YDynamics { get; set; }

        public ControlCellsQuery()
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

        public override Shapes.Controls.ControlCells CellDataToCellGroup(ShapeSheet.CellData[] row)
        {
            var cells = new Shapes.Controls.ControlCells();
            cells.CanGlue = row[this.CanGlue].ToInt();
            cells.Tip = row[this.Tip].ToInt();
            cells.X = row[this.X];
            cells.Y = row[this.Y];
            cells.YBehavior = row[this.YBehavior].ToInt();
            cells.XBehavior = row[this.XBehavior].ToInt();
            cells.XDynamics = row[this.XDynamics].ToInt();
            cells.YDynamics = row[this.YDynamics].ToInt();
            return cells;
        }
    }
}