using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using VACG = VisioAutomation.ShapeSheet.CellGroups;
using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.Shapes
{
    public class ConnectionPointCells : VACG.CellGroup
    {
        public Core.CellValue X { get; set; }
        public Core.CellValue Y { get; set; }
        public Core.CellValue DirX { get; set; }
        public Core.CellValue DirY { get; set; }
        public Core.CellValue Type { get; set; }

        public override IEnumerable<VACG.CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.X), Core.SrcConstants.ConnectionPointX, this.X);
            yield return this._create(nameof(this.Y), Core.SrcConstants.ConnectionPointY, this.Y);
            yield return this._create(nameof(this.DirX), Core.SrcConstants.ConnectionPointDirX, this.DirX);
            yield return this._create(nameof(this.DirY), Core.SrcConstants.ConnectionPointDirY, this.DirY);
            yield return this._create(nameof(this.Type), Core.SrcConstants.ConnectionPointType, this.Type);
        }

        public static List<List<ConnectionPointCells>> GetCells(IVisio.Page page, Core.ShapeIDPairs shapeidpairs,
            Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsMultipleShapesMultipleRows(page, shapeidpairs, type);
        }

        public static List<ConnectionPointCells> GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleShapeMultipleRows(shape, type);
        }

        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();

        class Builder : VACG.CellGroupBuilder<ConnectionPointCells>
        {
            public Builder() : base(VACG.CellGroupBuilderType.MultiRow)
            {
            }

            public override ConnectionPointCells ToCellGroup(VASS.Data.DataRow<string> row, VASS.Data.ColumnCollection cols)
            {
                var cells = new ConnectionPointCells();
                var getcellvalue = queryrow_to_cellgroup(row, cols);

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