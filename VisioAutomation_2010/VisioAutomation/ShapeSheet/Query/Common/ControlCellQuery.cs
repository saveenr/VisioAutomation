namespace VisioAutomation.ShapeSheet.Query.Common
{
    class ControlCellQuery : CellQuery
    {
        public Query.CellColumn CanGlue { get; set; }
        public Query.CellColumn Tip { get; set; }
        public Query.CellColumn X { get; set; }
        public Query.CellColumn Y { get; set; }
        public Query.CellColumn YBehavior { get; set; }
        public Query.CellColumn XBehavior { get; set; }
        public Query.CellColumn XDynamics { get; set; }
        public Query.CellColumn YDynamics { get; set; }

        public ControlCellQuery()
        {
            var sec = this.AddSection(Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionControls);
            this.CanGlue = sec.AddCell(ShapeSheet.SRCConstants.Controls_CanGlue, "Controls_CanGlue");
            this.Tip = sec.AddCell(ShapeSheet.SRCConstants.Controls_Tip, "Controls_Tip");
            this.X = sec.AddCell(ShapeSheet.SRCConstants.Controls_X, "Controls_X");
            this.Y = sec.AddCell(ShapeSheet.SRCConstants.Controls_Y, "Controls_Y");
            this.YBehavior = sec.AddCell(ShapeSheet.SRCConstants.Controls_YCon, "Controls_YCon");
            this.XBehavior = sec.AddCell(ShapeSheet.SRCConstants.Controls_XCon, "Controls_XCon");
            this.XDynamics = sec.AddCell(ShapeSheet.SRCConstants.Controls_XDyn, "Controls_XDyn");
            this.YDynamics = sec.AddCell(ShapeSheet.SRCConstants.Controls_YDyn, "Controls_YDyn");
        }

        public VisioAutomation.Shapes.Controls.ControlCells GetCells(System.Collections.Generic.IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VisioAutomation.Shapes.Controls.ControlCells();
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