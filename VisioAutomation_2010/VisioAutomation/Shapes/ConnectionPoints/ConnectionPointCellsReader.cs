using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.ConnectionPoints
{
    class ConnectionPointCellsReader : MultiRowReader<ConnectionPointCells>
    {
        public SubQueryColumn DirX { get; set; }
        public SubQueryColumn DirY { get; set; }
        public SubQueryColumn Type { get; set; }
        public SubQueryColumn X { get; set; }
        public SubQueryColumn Y { get; set; }

        public ConnectionPointCellsReader()
        {
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionConnectionPts);

            this.DirX = sec.AddCell(SrcConstants.Connections_DirX, nameof(SrcConstants.Connections_DirX));
            this.DirY = sec.AddCell(SrcConstants.Connections_DirY, nameof(SrcConstants.Connections_DirY));
            this.Type = sec.AddCell(SrcConstants.Connections_Type, nameof(SrcConstants.Connections_Type));
            this.X = sec.AddCell(SrcConstants.Connections_X, nameof(SrcConstants.Connections_X));
            this.Y = sec.AddCell(SrcConstants.Connections_Y, nameof(SrcConstants.Connections_Y));

        }

        public override ConnectionPointCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new ConnectionPointCells();
            cells.X = row[this.X];
            cells.Y = row[this.Y];
            cells.DirX = row[this.DirX];
            cells.DirY = row[this.DirY];
            cells.Type = row[this.Type];

            return cells;
        }
    }
}