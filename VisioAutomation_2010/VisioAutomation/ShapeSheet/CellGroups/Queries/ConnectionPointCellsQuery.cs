using VisioAutomation.Shapes.ConnectionPoints;
using VisioAutomation.ShapeSheet.Queries.Columns;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.ShapeSheet.CellGroups.Queries
{
    class ConnectionPointCellsQuery : CellGroupMultiRowQuery<ConnectionPointCells, double>
    {
        public ColumnSubQuery DirX { get; set; }
        public ColumnSubQuery DirY { get; set; }
        public ColumnSubQuery Type { get; set; }
        public ColumnSubQuery X { get; set; }
        public ColumnSubQuery Y { get; set; }

        public ConnectionPointCellsQuery()
        {
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionConnectionPts);

            this.DirX = sec.AddCell(SRCCON.Connections_DirX, nameof(SRCCON.Connections_DirX));
            this.DirY = sec.AddCell(SRCCON.Connections_DirY, nameof(SRCCON.Connections_DirY));
            this.Type = sec.AddCell(SRCCON.Connections_Type, nameof(SRCCON.Connections_Type));
            this.X = sec.AddCell(SRCCON.Connections_X, nameof(SRCCON.Connections_X));
            this.Y = sec.AddCell(SRCCON.Connections_Y, nameof(SRCCON.Connections_Y));

        }

        public override ConnectionPointCells CellDataToCellGroup(ShapeSheet.CellData[] row)
        {
            var cells = new ConnectionPointCells();
            cells.X = row[this.X];
            cells.Y = row[this.Y];
            cells.DirX = row[this.DirX].ToInt();
            cells.DirY = row[this.DirY].ToInt();
            cells.Type = row[this.Type].ToInt();

            return cells;
        }
    }
}