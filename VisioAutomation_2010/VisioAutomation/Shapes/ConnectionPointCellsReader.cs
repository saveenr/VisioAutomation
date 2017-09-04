using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    class ConnectionPointCellsReader : ReaderMultiRow<ConnectionPointCells>
    {
        public SectionQueryColumn DirX { get; set; }
        public SectionQueryColumn DirY { get; set; }
        public SectionQueryColumn Type { get; set; }
        public SectionQueryColumn X { get; set; }
        public SectionQueryColumn Y { get; set; }

        public ConnectionPointCellsReader()
        {
            var sec = this.query.SectionQueries.Add(IVisio.VisSectionIndices.visSectionConnectionPts);

            this.DirX = sec.Columns.Add(SrcConstants.ConnectionPointDirX, nameof(SrcConstants.ConnectionPointDirX));
            this.DirY = sec.Columns.Add(SrcConstants.ConnectionPointDirY, nameof(SrcConstants.ConnectionPointDirY));
            this.Type = sec.Columns.Add(SrcConstants.ConnectionPointType, nameof(SrcConstants.ConnectionPointType));
            this.X = sec.Columns.Add(SrcConstants.ConnectionPointX, nameof(SrcConstants.ConnectionPointX));
            this.Y = sec.Columns.Add(SrcConstants.ConnectionPointY, nameof(SrcConstants.ConnectionPointY));

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