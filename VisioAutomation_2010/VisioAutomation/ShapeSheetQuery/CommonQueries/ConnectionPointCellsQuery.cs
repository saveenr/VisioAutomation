using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheetQuery.CommonQueries
{
    class ConnectionPointCellsQuery : CellQuery
    {
        public CellColumn DirX { get; set; }
        public CellColumn DirY { get; set; }
        public CellColumn Type { get; set; }
        public CellColumn X { get; set; }
        public CellColumn Y { get; set; }

        public ConnectionPointCellsQuery()
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionConnectionPts);

            this.DirX = sec.AddCell(SRCCON.Connections_DirX, nameof(SRCCON.Connections_DirX));
            this.DirY = sec.AddCell(SRCCON.Connections_DirY, nameof(SRCCON.Connections_DirY));
            this.Type = sec.AddCell(SRCCON.Connections_Type, nameof(SRCCON.Connections_Type));
            this.X = sec.AddCell(SRCCON.Connections_X, nameof(SRCCON.Connections_X));
            this.Y = sec.AddCell(SRCCON.Connections_Y, nameof(SRCCON.Connections_Y));

        }

        public Shapes.Connections.ConnectionPointCells GetCells(ShapeSheet.CellData<double>[] row)
        {
            var cells = new Shapes.Connections.ConnectionPointCells();
            cells.X = row[this.X];
            cells.Y = row[this.Y];
            cells.DirX = Extensions.CellDataMethods.ToInt(row[this.DirX]);
            cells.DirY = Extensions.CellDataMethods.ToInt(row[this.DirY]);
            cells.Type = Extensions.CellDataMethods.ToInt(row[this.Type]);

            return cells;
        }
    }
}