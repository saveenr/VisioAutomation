using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellRecords;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class ControlCells : CellRecord
    {
        public Core.CellValue CanGlue { get; set; }
        public Core.CellValue Tip { get; set; }
        public Core.CellValue X { get; set; }
        public Core.CellValue Y { get; set; }
        public Core.CellValue YBehavior { get; set; }
        public Core.CellValue XBehavior { get; set; }
        public Core.CellValue XDynamics { get; set; }
        public Core.CellValue YDynamics { get; set; }

        public override IEnumerable<CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.CanGlue), Core.SrcConstants.ControlCanGlue, this.CanGlue);
            yield return this._create(nameof(this.Tip), Core.SrcConstants.ControlTip, this.Tip);
            yield return this._create(nameof(this.X), Core.SrcConstants.ControlX, this.X);
            yield return this._create(nameof(this.Y), Core.SrcConstants.ControlY, this.Y);
            yield return this._create(nameof(this.YBehavior), Core.SrcConstants.ControlYBehavior, this.YBehavior);
            yield return this._create(nameof(this.XBehavior), Core.SrcConstants.ControlXBehavior, this.XBehavior);
            yield return this._create(nameof(this.XDynamics), Core.SrcConstants.ControlXDynamics, this.XDynamics);
            yield return this._create(nameof(this.YDynamics), Core.SrcConstants.ControlYDynamics, this.YDynamics);
        }

        public static CellRecords<ControlCells> GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleShapeMultipleRows(shape, type);
        }

        public static CellRecordsGroup<ControlCells> GetCells(IVisio.Page page, Core.ShapeIDPairs shapeidpairs,
            Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsMultipleShapesMultipleRows(page, shapeidpairs, type);
        }


        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();

        public static ControlCells RowToRecord(VASS.Data.DataRow<string> row, VASS.Data.DataColumns cols)
        {
            var cells = new ControlCells();
            var getcellvalue = queryrow_to_cellrecord(row, cols);

            cells.CanGlue = getcellvalue(nameof(CanGlue));
            cells.Tip = getcellvalue(nameof(Tip));
            cells.X = getcellvalue(nameof(X));
            cells.Y = getcellvalue(nameof(Y));
            cells.YBehavior = getcellvalue(nameof(YBehavior));
            cells.XBehavior = getcellvalue(nameof(XBehavior));
            cells.XDynamics = getcellvalue(nameof(XDynamics));
            cells.YDynamics = getcellvalue(nameof(YDynamics));
            return cells;
        }

        class Builder : CellRecordBuilder<ControlCells>
        {
            public Builder() : base(CellRecordQueryType.SectionQuery, ControlCells.RowToRecord)
            {
            }

        }
    }
}