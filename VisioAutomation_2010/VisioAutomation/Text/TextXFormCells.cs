using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellRecords;
using VisioAutomation.ShapeSheet.Data;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class TextXFormCells : CellRecord
    {
        public Core.CellValue Angle { get; set; }
        public Core.CellValue Width { get; set; }
        public Core.CellValue Height { get; set; }
        public Core.CellValue PinX { get; set; }
        public Core.CellValue PinY { get; set; }
        public Core.CellValue LocPinX { get; set; }
        public Core.CellValue LocPinY { get; set; }

        public override IEnumerable<CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.PinX), Core.SrcConstants.TextXFormPinX, this.PinX);
            yield return this._create(nameof(this.PinY), Core.SrcConstants.TextXFormPinY, this.PinY);
            yield return this._create(nameof(this.LocPinX), Core.SrcConstants.TextXFormLocPinX, this.LocPinX);
            yield return this._create(nameof(this.LocPinY), Core.SrcConstants.TextXFormLocPinY, this.LocPinY);
            yield return this._create(nameof(this.Width), Core.SrcConstants.TextXFormWidth, this.Width);
            yield return this._create(nameof(this.Height), Core.SrcConstants.TextXFormHeight, this.Height);
            yield return this._create(nameof(this.Angle), Core.SrcConstants.TextXFormAngle, this.Angle);
        }

        public static CellRecords<TextXFormCells> GetCells(IVisio.Page page, IList<int> shapeids, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsMultipleShapesSingleRow(page, shapeids, type);
        }

        public static TextXFormCells GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleShapeSingleRow(shape, type);
        }

        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();

        public static TextXFormCells RowToRecord(DataRow<string> row, DataColumns cols)
        {
            var record = new TextXFormCells();
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


        class Builder : CellRecordBuilder<TextXFormCells>
        {
            public Builder() : base(CellRecordQueryType.CellQuery, TextXFormCells.RowToRecord)
            {
            }
        }


    }
}