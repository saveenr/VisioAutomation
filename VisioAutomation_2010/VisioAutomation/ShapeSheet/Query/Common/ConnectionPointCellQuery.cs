namespace VisioAutomation.ShapeSheet.Query.Common
{
    class ConnectionPointCellQuery : CellQuery
    {
        public Query.CellColumn DirX { get; set; }
        public Query.CellColumn DirY { get; set; }
        public Query.CellColumn Type { get; set; }
        public Query.CellColumn X { get; set; }
        public Query.CellColumn Y { get; set; }

        public ConnectionPointCellQuery()
        {
            var sec = this.AddSection(Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionConnectionPts);
            this.DirX = sec.AddCell(ShapeSheet.SRCConstants.Connections_DirX, "Connections_DirX");
            this.DirY = sec.AddCell(ShapeSheet.SRCConstants.Connections_DirY, "Connections_DirY");
            this.Type = sec.AddCell(ShapeSheet.SRCConstants.Connections_Type, "Connections_Type");
            this.X = sec.AddCell(ShapeSheet.SRCConstants.Connections_X, "Connections_X");
            this.Y = sec.AddCell(ShapeSheet.SRCConstants.Connections_Y, "Connections_Y");
        }

        public VisioAutomation.Shapes.Connections.ConnectionPointCells GetCells(System.Collections.Generic.IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VisioAutomation.Shapes.Connections.ConnectionPointCells();
            cells.X = row[this.X];
            cells.Y = row[this.Y];
            cells.DirX = Extensions.CellDataMethods.ToInt(row[this.DirX]);
            cells.DirY = Extensions.CellDataMethods.ToInt(row[this.DirY]);
            cells.Type = Extensions.CellDataMethods.ToInt(row[this.Type]);

            return cells;
        }
    }
}