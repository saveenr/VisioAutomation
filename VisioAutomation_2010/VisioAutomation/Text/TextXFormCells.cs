using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class TextXFormCells : VASS.CellGroups.CellGroup
    {
        public VisioAutomation.Core.CellValue Angle { get; set; }
        public VisioAutomation.Core.CellValue Width { get; set; }
        public VisioAutomation.Core.CellValue Height { get; set; }
        public VisioAutomation.Core.CellValue PinX { get; set; }
        public VisioAutomation.Core.CellValue PinY { get; set; }
        public VisioAutomation.Core.CellValue LocPinX { get; set; }
        public VisioAutomation.Core.CellValue LocPinY { get; set; }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.PinX), VisioAutomation.Core.SrcConstants.TextXFormPinX, this.PinX);
            yield return this.Create(nameof(this.PinY), VisioAutomation.Core.SrcConstants.TextXFormPinY, this.PinY);
            yield return this.Create(nameof(this.LocPinX), VisioAutomation.Core.SrcConstants.TextXFormLocPinX, this.LocPinX);
            yield return this.Create(nameof(this.LocPinY), VisioAutomation.Core.SrcConstants.TextXFormLocPinY, this.LocPinY);
            yield return this.Create(nameof(this.Width), VisioAutomation.Core.SrcConstants.TextXFormWidth, this.Width);
            yield return this.Create(nameof(this.Height), VisioAutomation.Core.SrcConstants.TextXFormHeight, this.Height);
            yield return this.Create(nameof(this.Angle), VisioAutomation.Core.SrcConstants.TextXFormAngle, this.Angle);
        }

        public static List<TextXFormCells> GetCells(IVisio.Page page, IList<int> shapeids, VisioAutomation.Core.CellValueType type)
        {
            var reader = TextXFormCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static TextXFormCells GetCells(IVisio.Shape shape, VisioAutomation.Core.CellValueType type)
        {
            var reader = TextXFormCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<TextXFormCellsBuilder> TextXFormCells_lazy_builder = new System.Lazy<TextXFormCellsBuilder>();


        class TextXFormCellsBuilder : VASS.CellGroups.CellGroupBuilder<Text.TextXFormCells>
        {
            public TextXFormCellsBuilder() : base(VisioAutomation.ShapeSheet.CellGroups.CellGroupBuilderType.SingleRow)
            {
            }

            public override Text.TextXFormCells ToCellGroup(ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols)
            {
                var cells = new Text.TextXFormCells();
                var getcellvalue = VisioAutomation.ShapeSheet.CellGroups.CellGroup.row_to_cellgroup(row, cols);

                cells.PinX = getcellvalue(nameof(TextXFormCells.PinX));
                cells.PinY = getcellvalue(nameof(TextXFormCells.PinY));
                cells.LocPinX = getcellvalue(nameof(TextXFormCells.LocPinX));
                cells.LocPinY = getcellvalue(nameof(TextXFormCells.LocPinY));
                cells.Width = getcellvalue(nameof(TextXFormCells.Width));
                cells.Height = getcellvalue(nameof(TextXFormCells.Height));
                cells.Angle = getcellvalue(nameof(TextXFormCells.Angle));

                return cells;
            }
        }


    }
}