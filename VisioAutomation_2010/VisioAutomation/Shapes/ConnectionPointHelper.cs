using VisioAutomation.ShapeSheet.CellGroups;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public static class ConnectionPointHelper
    {
        public static int Add(
            IVisio.Shape shape,
            ConnectionPointCells connection_point_cells)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            if (connection_point_cells.X.Value==null)
            {
                string msg = "Must provide an X Formula";
                throw new System.ArgumentException(msg, nameof(connection_point_cells));
            }

            if (connection_point_cells.Y.Value==null)
            {
                string msg = "Must provide an Y Formula";
                throw new System.ArgumentException(msg, nameof(connection_point_cells));
            }

            var n = shape.AddRow((short)IVisio.VisSectionIndices.visSectionConnectionPts,
                                 (short)IVisio.VisRowIndices.visRowLast,
                                 (short)IVisio.VisRowTags.visTagCnnctPt);

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetFormulas(connection_point_cells, n);

            writer.Commit(shape);

            return n;
        }

        public static void Delete(IVisio.Shape shape, int index)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            if (index < 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(index));
            }

            var row = (IVisio.VisRowIndices)index;
            shape.DeleteRow( (short) IVisio.VisSectionIndices.visSectionConnectionPts, (short)row);
        }

        public static int GetCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            return shape.RowCount[ (short) IVisio.VisSectionIndices.visSectionConnectionPts];
        }

        public static int Delete(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            int n = ConnectionPointHelper.GetCount(shape);
            for (int i = n - 1; i >= 0; i--)
            {
                ConnectionPointHelper.Delete(shape, i);
            }

            return n;
        }

        public static List<List<ConnectionPointCells>> GetConnectionPointCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = ConnectionPointCells_lazy_reader.Value;
            return reader.GetCellsMultiRow(page, shapeids, type);
        }

        public static List<ConnectionPointCells> GetConnectionPointCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = ConnectionPointCells_lazy_reader.Value;
            return reader.GetCellsMultiRow(shape, type);
        }

        private static readonly System.Lazy<ConnectionPointCellsReader> ConnectionPointCells_lazy_reader = new System.Lazy<ConnectionPointCellsReader>();

        class ConnectionPointCellsReader : CellGroupReader<ConnectionPointCells>
        {
            public SectionQueryColumn DirX { get; set; }
            public SectionQueryColumn DirY { get; set; }
            public SectionQueryColumn Type { get; set; }
            public SectionQueryColumn X { get; set; }
            public SectionQueryColumn Y { get; set; }

            public ConnectionPointCellsReader()
                : base(new VisioAutomation.ShapeSheet.Query.SectionsQuery())
            {
                var sec = this.query_multirow.SectionQueries.Add(IVisio.VisSectionIndices.visSectionConnectionPts);

                this.DirX = sec.Columns.Add(SrcConstants.ConnectionPointDirX, nameof(this.DirX));
                this.DirY = sec.Columns.Add(SrcConstants.ConnectionPointDirY, nameof(this.DirY));
                this.Type = sec.Columns.Add(SrcConstants.ConnectionPointType, nameof(this.Type));
                this.X = sec.Columns.Add(SrcConstants.ConnectionPointX, nameof(this.X));
                this.Y = sec.Columns.Add(SrcConstants.ConnectionPointY, nameof(this.Y));

            }

            public override ConnectionPointCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
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