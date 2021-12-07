using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellRecords;
using VACG=VisioAutomation.ShapeSheet.CellGroups;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class TextXFormCells : VACG.CellRecord
    {
        public Core.CellValue Angle { get; set; }
        public Core.CellValue Width { get; set; }
        public Core.CellValue Height { get; set; }
        public Core.CellValue PinX { get; set; }
        public Core.CellValue PinY { get; set; }
        public Core.CellValue LocPinX { get; set; }
        public Core.CellValue LocPinY { get; set; }

        public override IEnumerable<VACG.CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.PinX), Core.SrcConstants.TextXFormPinX, this.PinX);
            yield return this._create(nameof(this.PinY), Core.SrcConstants.TextXFormPinY, this.PinY);
            yield return this._create(nameof(this.LocPinX), Core.SrcConstants.TextXFormLocPinX, this.LocPinX);
            yield return this._create(nameof(this.LocPinY), Core.SrcConstants.TextXFormLocPinY, this.LocPinY);
            yield return this._create(nameof(this.Width), Core.SrcConstants.TextXFormWidth, this.Width);
            yield return this._create(nameof(this.Height), Core.SrcConstants.TextXFormHeight, this.Height);
            yield return this._create(nameof(this.Angle), Core.SrcConstants.TextXFormAngle, this.Angle);
        }

        public static List<TextXFormCells> GetCells(IVisio.Page page, IList<int> shapeids, Core.CellValueType type)
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


        class Builder : CellGroupBuilder<TextXFormCells>
        {
            public Builder() : base(VACG.CellGroupBuilderType.SingleRow)
            {
            }

            public override TextXFormCells ToCellGroup(VASS.Data.DataRow<string> row, VASS.Data.DataColumnCollection cols)
            {
                var cells = new TextXFormCells();
                var getcellvalue = queryrow_to_cellgroup(row, cols);

                cells.PinX = getcellvalue(nameof(PinX));
                cells.PinY = getcellvalue(nameof(PinY));
                cells.LocPinX = getcellvalue(nameof(LocPinX));
                cells.LocPinY = getcellvalue(nameof(LocPinY));
                cells.Width = getcellvalue(nameof(Width));
                cells.Height = getcellvalue(nameof(Height));
                cells.Angle = getcellvalue(nameof(Angle));

                return cells;
            }
        }


    }
}