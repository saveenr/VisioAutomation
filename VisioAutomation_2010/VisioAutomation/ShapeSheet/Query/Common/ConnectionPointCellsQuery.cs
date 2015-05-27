using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.ShapeSheet.Query.Common
{
    class ConnectionPointCellsQuery : CellQuery
    {
        public Query.CellColumn DirX { get; set; }
        public Query.CellColumn DirY { get; set; }
        public Query.CellColumn Type { get; set; }
        public Query.CellColumn X { get; set; }
        public Query.CellColumn Y { get; set; }

        public ConnectionPointCellsQuery()
        {
            var sec = this.AddSection(Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionConnectionPts);

            this.DirX = sec.AddCell(SRCCON.Connections_DirX, nameof(SRCCON.Connections_DirX));
            this.DirY = sec.AddCell(SRCCON.Connections_DirY, nameof(SRCCON.Connections_DirY));
            this.Type = sec.AddCell(SRCCON.Connections_Type, nameof(SRCCON.Connections_Type));
            this.X = sec.AddCell(SRCCON.Connections_X, nameof(SRCCON.Connections_X));
            this.Y = sec.AddCell(SRCCON.Connections_Y, nameof(SRCCON.Connections_Y));

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