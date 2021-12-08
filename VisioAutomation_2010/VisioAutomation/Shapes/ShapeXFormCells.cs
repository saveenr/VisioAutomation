using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellRecords;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioAutomation.Shapes
{
    public class ShapeXFormCells : CellRecord
    {
        public Core.CellValue PinX { get; set; }
        public Core.CellValue PinY { get; set; }
        public Core.CellValue LocPinX { get; set; }
        public Core.CellValue LocPinY { get; set; }
        public Core.CellValue Width { get; set; }
        public Core.CellValue Height { get; set; }
        public Core.CellValue Angle { get; set; }

        public override IEnumerable<CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.PinX), Core.SrcConstants.XFormPinX, this.PinX);
            yield return this._create(nameof(this.PinY), Core.SrcConstants.XFormPinY, this.PinY);
            yield return this._create(nameof(this.LocPinX), Core.SrcConstants.XFormLocPinX, this.LocPinX);
            yield return this._create(nameof(this.LocPinY), Core.SrcConstants.XFormLocPinY, this.LocPinY);
            yield return this._create(nameof(this.Width), Core.SrcConstants.XFormWidth, this.Width);
            yield return this._create(nameof(this.Height), Core.SrcConstants.XFormHeight, this.Height);
            yield return this._create(nameof(this.Angle), Core.SrcConstants.XFormAngle, this.Angle);
        }


        public static List<ShapeXFormCells> GetCells(IVisio.Page page, IList<int> shapeids, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsMultipleShapesSingleRow(page, shapeids, type);
        }

        public static ShapeXFormCells GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleShapeSingleRow(shape, type);
        }

        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();

        public static ShapeXFormCells RowToRecord(VASS.Data.DataRow<string> row, VASS.Data.DataColumns cols)
        {
            var record = new ShapeXFormCells();
            var getcellvalue = queryrow_to_cellrecord(row, cols);

            record.PinX = getcellvalue(nameof(PinX));
            record.PinY = getcellvalue(nameof(PinY));
            record.LocPinX = getcellvalue(nameof(LocPinX));
            record.LocPinY = getcellvalue(nameof(LocPinY));
            record.Width = getcellvalue(nameof(Width));
            record.Height = getcellvalue(nameof(Height));
            record.Angle = getcellvalue(nameof(Angle));

            return record;
        }

        class Builder : CellRecordBuilder<ShapeXFormCells>
        {
            public Builder() : base(CellRecordQueryType.CellQuery, ShapeXFormCells.RowToRecord)
            {
            }

        }
    }
}