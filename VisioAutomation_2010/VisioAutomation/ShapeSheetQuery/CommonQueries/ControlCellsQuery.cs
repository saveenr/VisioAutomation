using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheetQuery.CommonQueries
{
    class ControlCellsQuery : CellQuery
    {
        public CellColumn CanGlue { get; set; }
        public CellColumn Tip { get; set; }
        public CellColumn X { get; set; }
        public CellColumn Y { get; set; }
        public CellColumn YBehavior { get; set; }
        public CellColumn XBehavior { get; set; }
        public CellColumn XDynamics { get; set; }
        public CellColumn YDynamics { get; set; }

        public ControlCellsQuery()
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionControls);

            this.CanGlue = sec.AddCell(SRCCON.Controls_CanGlue, nameof(SRCCON.Controls_CanGlue));
            this.Tip = sec.AddCell(SRCCON.Controls_Tip, nameof(SRCCON.Controls_Tip));
            this.X = sec.AddCell(SRCCON.Controls_X, nameof(SRCCON.Controls_X));
            this.Y = sec.AddCell(SRCCON.Controls_Y, nameof(SRCCON.Controls_Y));
            this.YBehavior = sec.AddCell(SRCCON.Controls_YCon, nameof(SRCCON.Controls_YCon));
            this.XBehavior = sec.AddCell(SRCCON.Controls_XCon, nameof(SRCCON.Controls_XCon));
            this.XDynamics = sec.AddCell(SRCCON.Controls_XDyn, nameof(SRCCON.Controls_XDyn));
            this.YDynamics = sec.AddCell(SRCCON.Controls_YDyn, nameof(SRCCON.Controls_YDyn));

        }

        public Shapes.Controls.ControlCells GetCells(ShapeSheet.CellData<double>[] row)
        {
            var cells = new Shapes.Controls.ControlCells();
            cells.CanGlue = Extensions.CellDataMethods.ToInt(row[this.CanGlue]);
            cells.Tip = Extensions.CellDataMethods.ToInt(row[this.Tip]);
            cells.X = row[this.X];
            cells.Y = row[this.Y];
            cells.YBehavior = Extensions.CellDataMethods.ToInt(row[this.YBehavior]);
            cells.XBehavior = Extensions.CellDataMethods.ToInt(row[this.XBehavior]);
            cells.XDynamics = Extensions.CellDataMethods.ToInt(row[this.XDynamics]);
            cells.YDynamics = Extensions.CellDataMethods.ToInt(row[this.YDynamics]);
            return cells;
        }
    }
}