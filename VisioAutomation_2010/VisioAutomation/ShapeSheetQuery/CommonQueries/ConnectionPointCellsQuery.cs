using VisioAutomation.Shapes.ConnectionPoints;
using VisioAutomation.ShapeSheetQuery.Columns;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheetQuery.CommonQueries
{
    class ConnectionPointCellsQuery : Query
    {
        public ColumnSubQuery DirX { get; set; }
        public ColumnSubQuery DirY { get; set; }
        public ColumnSubQuery Type { get; set; }
        public ColumnSubQuery X { get; set; }
        public ColumnSubQuery Y { get; set; }

        public ConnectionPointCellsQuery()
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionConnectionPts);

            this.DirX = sec.AddCell(SRCCON.Connections_DirX, nameof(SRCCON.Connections_DirX));
            this.DirY = sec.AddCell(SRCCON.Connections_DirY, nameof(SRCCON.Connections_DirY));
            this.Type = sec.AddCell(SRCCON.Connections_Type, nameof(SRCCON.Connections_Type));
            this.X = sec.AddCell(SRCCON.Connections_X, nameof(SRCCON.Connections_X));
            this.Y = sec.AddCell(SRCCON.Connections_Y, nameof(SRCCON.Connections_Y));

        }

        public ConnectionPointCells GetCells(ShapeSheet.CellData<double>[] row)
        {
            var cells = new ConnectionPointCells();
            cells.X = row[this.X];
            cells.Y = row[this.Y];
            cells.DirX = Extensions.CellDataMethods.ToInt(row[this.DirX]);
            cells.DirY = Extensions.CellDataMethods.ToInt(row[this.DirY]);
            cells.Type = Extensions.CellDataMethods.ToInt(row[this.Type]);

            return cells;
        }
    }
}