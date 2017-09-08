using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class ConnectionPointCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral X { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Y { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral DirX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral DirY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Type { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ConnectionPointX, this.X.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ConnectionPointY, this.Y.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ConnectionPointDirX, this.DirX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ConnectionPointDirY, this.DirY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ConnectionPointType, this.Type.Value);
            }
        }

        public static List<List<ConnectionPointCells>> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var query = ConnectionPointCells.lazy_query.Value;
            return query.GetValues(page, shapeids, CellValueType.Formula);
        }

        public static List<List<ConnectionPointCells>> GetResults(IVisio.Page page, IList<int> shapeids)
        {
            var query = ConnectionPointCells.lazy_query.Value;
            return query.GetValues(page, shapeids, CellValueType.Result);
        }

        public static List<ConnectionPointCells> GetFormulas(IVisio.Shape shape)
        {
            var query = ConnectionPointCells.lazy_query.Value;
            return query.GetValues(shape, CellValueType.Formula);
        }

        public static List<ConnectionPointCells> GetResults(IVisio.Shape shape)
        {
            var query = ConnectionPointCells.lazy_query.Value;
            return query.GetValues(shape, CellValueType.Result);
        }

        private static readonly System.Lazy<ConnectionPointCellsReader> lazy_query = new System.Lazy<ConnectionPointCellsReader>();

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

            public override ConnectionPointCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row)
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
}