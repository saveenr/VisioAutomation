using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    class ConnectionPointCellsReader : ReaderMultiRow<ConnectionPointCells>
    {
        public SubQueryColumn DirX { get; set; }
        public SubQueryColumn DirY { get; set; }
        public SubQueryColumn Type { get; set; }
        public SubQueryColumn X { get; set; }
        public SubQueryColumn Y { get; set; }

        public ConnectionPointCellsReader()
        {
            var sec = this.query.AddSubQuery(IVisio.VisSectionIndices.visSectionConnectionPts);

            this.DirX = sec.AddCell(SrcConstants.ConnectionPointDirX, nameof(SrcConstants.ConnectionPointDirX));
            this.DirY = sec.AddCell(SrcConstants.ConnectionPointDirY, nameof(SrcConstants.ConnectionPointDirY));
            this.Type = sec.AddCell(SrcConstants.ConnectionPointType, nameof(SrcConstants.ConnectionPointType));
            this.X = sec.AddCell(SrcConstants.ConnectionPointX, nameof(SrcConstants.ConnectionPointX));
            this.Y = sec.AddCell(SrcConstants.ConnectionPointY, nameof(SrcConstants.ConnectionPointY));

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