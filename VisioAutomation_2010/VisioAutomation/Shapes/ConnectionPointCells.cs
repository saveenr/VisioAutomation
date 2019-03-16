using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Shapes
{
    public class ConnectionPointCells : CellGroup
    {
        public CellValueLiteral X { get; set; }
        public CellValueLiteral Y { get; set; }
        public CellValueLiteral DirX { get; set; }
        public CellValueLiteral DirY { get; set; }
        public CellValueLiteral Type { get; set; }

        public override IEnumerable<CellMetadataItem> CellMetadata
        {
            get
            {
                yield return CellMetadataItem.Create(nameof(this.X), SrcConstants.ConnectionPointX, this.X);
                yield return CellMetadataItem.Create(nameof(this.Y), SrcConstants.ConnectionPointY, this.Y);
                yield return CellMetadataItem.Create(nameof(this.DirX), SrcConstants.ConnectionPointDirX, this.DirX);
                yield return CellMetadataItem.Create(nameof(this.DirY), SrcConstants.ConnectionPointDirY, this.DirY);
                yield return CellMetadataItem.Create(nameof(this.Type), SrcConstants.ConnectionPointType, this.Type);
            }
        }

        public static List<List<ConnectionPointCells>> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = ConnectionPointCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(page, shapeids, type);
        }

        public static List<ConnectionPointCells> GetCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = ConnectionPointCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(shape, type);
        }

        private static readonly System.Lazy<ConnectionPointCellsBuilder> ConnectionPointCells_lazy_builder = new System.Lazy<ConnectionPointCellsBuilder>();

        class ConnectionPointCellsBuilder : CellGroupBuilder<ConnectionPointCells>
        {

            public ConnectionPointCellsBuilder() : base(CellGroupBuilderType.MultiRow)
            {
            }

            public override ConnectionPointCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row, VisioAutomation.ShapeSheet.Query.ColumnList cols)
            {
                var cells = new ConnectionPointCells();

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }

                cells.X = getcellvalue(nameof(ConnectionPointCells.X));
                cells.Y = getcellvalue(nameof(ConnectionPointCells.Y));
                cells.DirX = getcellvalue(nameof(ConnectionPointCells.DirX));
                cells.DirY = getcellvalue(nameof(ConnectionPointCells.DirY));
                cells.Type = getcellvalue(nameof(ConnectionPointCells.Type));

                return cells;
            }
        }

    }
}