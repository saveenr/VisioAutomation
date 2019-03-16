using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class TextXFormCells : CellGroup
    {
        public CellValueLiteral Angle { get; set; }
        public CellValueLiteral Width { get; set; }
        public CellValueLiteral Height { get; set; }
        public CellValueLiteral PinX { get; set; }
        public CellValueLiteral PinY { get; set; }
        public CellValueLiteral LocPinX { get; set; }
        public CellValueLiteral LocPinY { get; set; }

        public override IEnumerable<CellMetadataItem> CellMetadata
        {
            get
            {


                yield return CellMetadataItem.Create(nameof(this.PinX), SrcConstants.TextXFormPinX, this.PinX);
                yield return CellMetadataItem.Create(nameof(this.PinY), SrcConstants.TextXFormPinY, this.PinY);
                yield return CellMetadataItem.Create(nameof(this.LocPinX), SrcConstants.TextXFormLocPinX, this.LocPinX);
                yield return CellMetadataItem.Create(nameof(this.LocPinY), SrcConstants.TextXFormLocPinY, this.LocPinY);
                yield return CellMetadataItem.Create(nameof(this.Width), SrcConstants.TextXFormWidth, this.Width);
                yield return CellMetadataItem.Create(nameof(this.Height), SrcConstants.TextXFormHeight, this.Height);
                yield return CellMetadataItem.Create(nameof(this.Angle), SrcConstants.TextXFormAngle, this.Angle);
            }
        }

        public static List<TextXFormCells> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = TextXFormCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static TextXFormCells GetCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = TextXFormCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<TextXFormCellsBuilder> TextXFormCells_lazy_builder = new System.Lazy<TextXFormCellsBuilder>();


        class TextXFormCellsBuilder : CellGroupBuilder<Text.TextXFormCells>
        {
            public TextXFormCellsBuilder() : base(VisioAutomation.ShapeSheet.CellGroups.CellGroupBuilderType.SingleRow)
            {
            }

            public override Text.TextXFormCells ToCellGroup(VisioAutomation.ShapeSheet.Internal.ArraySegment<string> row, VisioAutomation.ShapeSheet.Query.ColumnList cols)
            {
                var cells = new Text.TextXFormCells();

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }

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