using VisioAutomation.ShapeSheet.CellGroups;


namespace VisioAutomation.Pages;

public class PageRulerAndGridCells : VASS.CellGroups.CellGroup
{
    public VASS.CellValue XGridDensity { get; set; }
    public VASS.CellValue YGridDensity { get; set; }
    public VASS.CellValue XGridOrigin { get; set; }
    public VASS.CellValue YGridOrigin { get; set; }
    public VASS.CellValue XGridSpacing { get; set; }
    public VASS.CellValue YGridSpacing { get; set; }
    public VASS.CellValue XRulerDensity { get; set; }
    public VASS.CellValue XRulerOrigin { get; set; }
    public VASS.CellValue YRulerDensity { get; set; }
    public VASS.CellValue YRulerOrigin { get; set; }

    public override IEnumerable<CellMetadataItem> GetCellMetadata()
    {
        yield return this.Create(nameof(this.XGridDensity), VASS.SrcConstants.XGridDensity, this.XGridDensity);
        yield return this.Create(nameof(this.XGridOrigin), VASS.SrcConstants.XGridOrigin, this.XGridOrigin);
        yield return this.Create(nameof(this.XGridSpacing), VASS.SrcConstants.XGridSpacing, this.XGridSpacing);
        yield return this.Create(nameof(this.XRulerDensity), VASS.SrcConstants.XRulerDensity, this.XRulerDensity);
        yield return this.Create(nameof(this.XRulerOrigin), VASS.SrcConstants.XRulerOrigin, this.XRulerOrigin);
        yield return this.Create(nameof(this.YGridDensity), VASS.SrcConstants.YGridDensity, this.YGridDensity);
        yield return this.Create(nameof(this.YGridOrigin), VASS.SrcConstants.YGridOrigin, this.YGridOrigin);
        yield return this.Create(nameof(this.YGridSpacing), VASS.SrcConstants.YGridSpacing, this.YGridSpacing);
        yield return this.Create(nameof(this.YRulerDensity), VASS.SrcConstants.YRulerDensity, this.YRulerDensity);
        yield return this.Create(nameof(this.YRulerOrigin), VASS.SrcConstants.YRulerOrigin, this.YRulerOrigin);
    }

    public static PageRulerAndGridCells GetCells(IVisio.Shape shape, VASS.CellValueType type)
    {
        var reader = PageRulerAndGridCells_lazy_builder.Value;
        return reader.GetCellsSingleRow(shape, type);
    }

    private static readonly System.Lazy<PageRulerAndGridCellsBuilder> PageRulerAndGridCells_lazy_builder = new System.Lazy<PageRulerAndGridCellsBuilder>();

    class PageRulerAndGridCellsBuilder : VASS.CellGroups.CellGroupBuilder<PageRulerAndGridCells>
    {
        public PageRulerAndGridCellsBuilder() : base(VASS.CellGroups.CellGroupBuilderType.SingleRow)
        {
        }

        public override PageRulerAndGridCells ToCellGroup(ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols)
        {
            var cells = new PageRulerAndGridCells();
            var getcellvalue = VisioAutomation.ShapeSheet.CellGroups.CellGroup.row_to_cellgroup(row, cols);

            cells.XGridDensity = getcellvalue(nameof(PageRulerAndGridCells.XGridDensity));
            cells.XGridOrigin = getcellvalue(nameof(PageRulerAndGridCells.XGridOrigin));
            cells.XGridSpacing = getcellvalue(nameof(PageRulerAndGridCells.XGridSpacing));
            cells.XRulerDensity = getcellvalue(nameof(PageRulerAndGridCells.XRulerDensity));
            cells.XRulerOrigin = getcellvalue(nameof(PageRulerAndGridCells.XRulerOrigin));
            cells.YGridDensity = getcellvalue(nameof(PageRulerAndGridCells.YGridDensity));
            cells.YGridOrigin = getcellvalue(nameof(PageRulerAndGridCells.YGridOrigin));
            cells.YGridSpacing = getcellvalue(nameof(PageRulerAndGridCells.YGridSpacing));
            cells.YRulerDensity = getcellvalue(nameof(PageRulerAndGridCells.YRulerDensity));
            cells.YRulerOrigin = getcellvalue(nameof(PageRulerAndGridCells.YRulerOrigin));

            return cells;
        }
    }

}