using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellRecords;
using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.Shapes
{
    public class ConnectionPointCells : CellRecord
    {
        public Core.CellValue X { get; set; }
        public Core.CellValue Y { get; set; }
        public Core.CellValue DirX { get; set; }
        public Core.CellValue DirY { get; set; }
        public Core.CellValue Type { get; set; }

        public override IEnumerable<CellMetadata> GetCellMetadata()
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

        class Builder : CellRecordBuilder<ConnectionPointCells>
        {
            public Builder() : base(CellRecordQueryType.SectionQuery)
            {
            }

            public override ConnectionPointCells RowToRecord(VASS.Data.DataRow<string> row, VASS.Data.DataColumns cols)
            {
                var record = new ConnectionPointCells();
                var getcellvalue = queryrow_to_cellrecord(row, cols);

                record.X = getcellvalue(nameof(X));
                record.Y = getcellvalue(nameof(Y));
                record.DirX = getcellvalue(nameof(DirX));
                record.DirY = getcellvalue(nameof(DirY));
                record.Type = getcellvalue(nameof(ConnectionPointCells.Type));

                return record;
            }
        }
    }
}