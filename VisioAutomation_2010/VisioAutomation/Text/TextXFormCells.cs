
using VisioAutomation.ShapeSheet.CellGroups;


namespace VisioAutomation.Text;

public class TextXFormCells : VASS.CellGroups.CellGroup
{
    public VASS.CellValue Angle { get; set; }
    public VASS.CellValue Width { get; set; }
    public VASS.CellValue Height { get; set; }
    public VASS.CellValue PinX { get; set; }
    public VASS.CellValue PinY { get; set; }
    public VASS.CellValue LocPinX { get; set; }
    public VASS.CellValue LocPinY { get; set; }

    public override IEnumerable<CellMetadataItem> GetCellMetadata()
    {
        yield return this.Create(nameof(this.PinX), VASS.SrcConstants.TextXFormPinX, this.PinX);
        yield return this.Create(nameof(this.PinY), VASS.SrcConstants.TextXFormPinY, this.PinY);
        yield return this.Create(nameof(this.LocPinX), VASS.SrcConstants.TextXFormLocPinX, this.LocPinX);
        yield return this.Create(nameof(this.LocPinY), VASS.SrcConstants.TextXFormLocPinY, this.LocPinY);
        yield return this.Create(nameof(this.Width), VASS.SrcConstants.TextXFormWidth, this.Width);
        yield return this.Create(nameof(this.Height), VASS.SrcConstants.TextXFormHeight, this.Height);
        yield return this.Create(nameof(this.Angle), VASS.SrcConstants.TextXFormAngle, this.Angle);
    }

    public static List<TextXFormCells> GetCells(IVisio.Page page, IList<int> shapeids, VASS.CellValueType type)
    {
        var reader = TextXFormCells_lazy_builder.Value;
        return reader.GetCellsSingleRow(page, shapeids, type);
    }

    public static TextXFormCells GetCells(IVisio.Shape shape, VASS.CellValueType type)
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