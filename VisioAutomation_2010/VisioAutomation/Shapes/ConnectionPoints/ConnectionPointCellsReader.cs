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

            this.DirX = sec.AddCell(SrcConstants.ConnectionDirX, nameof(SrcConstants.ConnectionDirX));
            this.DirY = sec.AddCell(SrcConstants.ConnectionDirY, nameof(SrcConstants.ConnectionDirY));
            this.Type = sec.AddCell(SrcConstants.ConnectionType, nameof(SrcConstants.ConnectionType));
            this.X = sec.AddCell(SrcConstants.ConnectionX, nameof(SrcConstants.ConnectionX));
            this.Y = sec.AddCell(SrcConstants.ConnectionY, nameof(SrcConstants.ConnectionY));

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