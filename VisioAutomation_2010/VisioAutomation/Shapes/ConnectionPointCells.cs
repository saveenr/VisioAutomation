using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;

namespace VisioAutomation.Shapes
{
    public class ConnectionPointCells : CellGroup
    {
        public Core.CellValue X { get; set; }
        public Core.CellValue Y { get; set; }
        public Core.CellValue DirX { get; set; }
        public Core.CellValue DirY { get; set; }
        public Core.CellValue Type { get; set; }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.X), Core.SrcConstants.ConnectionPointX, this.X);
            yield return this.Create(nameof(this.Y), Core.SrcConstants.ConnectionPointY, this.Y);
            yield return this.Create(nameof(this.DirX), Core.SrcConstants.ConnectionPointDirX, this.DirX);
            yield return this.Create(nameof(this.DirY), Core.SrcConstants.ConnectionPointDirY, this.DirY);
            yield return this.Create(nameof(this.Type), Core.SrcConstants.ConnectionPointType, this.Type);
        }

        public static List<List<ConnectionPointCells>> GetCells(IVisio.Page page, Core.ShapeIDPairs shapeidpairs, Core.CellValueType type)
        {
            var reader = ConnectionPointCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(page, shapeidpairs, type);
        }

        public static List<ConnectionPointCells> GetCells(IVisio.Shape shape, Core.CellValueType type)
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

            public override ConnectionPointCells ToCellGroup(VASS.Query.Row<string> row, VASS.Query.Columns cols)
            {
                var cells = new ConnectionPointCells();
                var getcellvalue = row_to_cellgroup(row, cols);

                cells.X = getcellvalue(nameof(X));
                cells.Y = getcellvalue(nameof(Y));
                cells.DirX = getcellvalue(nameof(DirX));
                cells.DirY = getcellvalue(nameof(DirY));
                cells.Type = getcellvalue(nameof(ConnectionPointCells.Type));

                return cells;
            }
        }

    }
}