using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class TextXFormCells : CellGroup
    {
        public Core.CellValue Angle { get; set; }
        public Core.CellValue Width { get; set; }
        public Core.CellValue Height { get; set; }
        public Core.CellValue PinX { get; set; }
        public Core.CellValue PinY { get; set; }
        public Core.CellValue LocPinX { get; set; }
        public Core.CellValue LocPinY { get; set; }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.PinX), Core.SrcConstants.TextXFormPinX, this.PinX);
            yield return this.Create(nameof(this.PinY), Core.SrcConstants.TextXFormPinY, this.PinY);
            yield return this.Create(nameof(this.LocPinX), Core.SrcConstants.TextXFormLocPinX, this.LocPinX);
            yield return this.Create(nameof(this.LocPinY), Core.SrcConstants.TextXFormLocPinY, this.LocPinY);
            yield return this.Create(nameof(this.Width), Core.SrcConstants.TextXFormWidth, this.Width);
            yield return this.Create(nameof(this.Height), Core.SrcConstants.TextXFormHeight, this.Height);
            yield return this.Create(nameof(this.Angle), Core.SrcConstants.TextXFormAngle, this.Angle);
        }

        public static List<TextXFormCells> GetCells(IVisio.Page page, IList<int> shapeids, Core.CellValueType type)
        {
            var reader = TextXFormCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static TextXFormCells GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = TextXFormCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<Builder> TextXFormCells_lazy_builder = new System.Lazy<Builder>();


        class Builder : CellGroupBuilder<TextXFormCells>
        {
            public Builder() : base(CellGroupBuilderType.SingleRow)
            {
            }

            public override TextXFormCells ToCellGroup(VASS.Query.Row<string> row, VASS.Query.Columns cols)
            {
                var cells = new TextXFormCells();
                var getcellvalue = row_to_cellgroup(row, cols);

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