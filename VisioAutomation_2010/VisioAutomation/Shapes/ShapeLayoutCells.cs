using VisioAutomation.ShapeSheet.CellGroups;


namespace VisioAutomation.Shapes;

public class ShapeLayoutCells : VASS.CellGroups.CellGroup
{
    public VASS.CellValue ConnectorFixedCode { get; set; }
    public VASS.CellValue LineJumpCode { get; set; }
    public VASS.CellValue LineJumpDirX { get; set; }
    public VASS.CellValue LineJumpDirY { get; set; }
    public VASS.CellValue LineJumpStyle { get; set; }
    public VASS.CellValue LineRouteExt { get; set; }
    public VASS.CellValue ShapeFixedCode { get; set; }
    public VASS.CellValue ShapePermeablePlace { get; set; }
    public VASS.CellValue ShapePermeableX { get; set; }
    public VASS.CellValue ShapePermeableY { get; set; }
    public VASS.CellValue ShapePlaceFlip { get; set; }
    public VASS.CellValue ShapePlaceStyle { get; set; }
    public VASS.CellValue ShapePlowCode { get; set; }
    public VASS.CellValue ShapeRouteStyle { get; set; }
    public VASS.CellValue ShapeSplit { get; set; }
    public VASS.CellValue ShapeSplittable { get; set; }
    public VASS.CellValue ShapeDisplayLevel { get; set; } // new in visio 2010
    public VASS.CellValue Relationships { get; set; } // new in visio 2010

    public override IEnumerable<CellMetadataItem> GetCellMetadata()
    {
        yield return this.Create(nameof(this.ConnectorFixedCode), VASS.SrcConstants.ShapeLayoutConnectorFixedCode,
            this.ConnectorFixedCode);
        yield return this.Create(nameof(this.LineJumpCode), VASS.SrcConstants.ShapeLayoutLineJumpCode, this.LineJumpCode);
        yield return this.Create(nameof(this.LineJumpDirX), VASS.SrcConstants.ShapeLayoutLineJumpDirX, this.LineJumpDirX);
        yield return this.Create(nameof(this.LineJumpDirY), VASS.SrcConstants.ShapeLayoutLineJumpDirY, this.LineJumpDirY);
        yield return this.Create(nameof(this.LineJumpStyle), VASS.SrcConstants.ShapeLayoutLineJumpStyle,
            this.LineJumpStyle);
        yield return this.Create(nameof(this.LineRouteExt), VASS.SrcConstants.ShapeLayoutLineRouteExt, this.LineRouteExt);
        yield return this.Create(nameof(this.ShapeFixedCode), VASS.SrcConstants.ShapeLayoutShapeFixedCode,
            this.ShapeFixedCode);
        yield return this.Create(nameof(this.ShapePermeablePlace), VASS.SrcConstants.ShapeLayoutShapePermeablePlace,
            this.ShapePermeablePlace);
        yield return this.Create(nameof(this.ShapePermeableX), VASS.SrcConstants.ShapeLayoutShapePermeableX,
            this.ShapePermeableX);
        yield return this.Create(nameof(this.ShapePermeableY), VASS.SrcConstants.ShapeLayoutShapePermeableY,
            this.ShapePermeableY);
        yield return this.Create(nameof(this.ShapePlaceFlip), VASS.SrcConstants.ShapeLayoutShapePlaceFlip,
            this.ShapePlaceFlip);
        yield return this.Create(nameof(this.ShapePlaceStyle), VASS.SrcConstants.ShapeLayoutShapePlaceStyle,
            this.ShapePlaceStyle);
        yield return this.Create(nameof(this.ShapePlowCode), VASS.SrcConstants.ShapeLayoutShapePlowCode,
            this.ShapePlowCode);
        yield return this.Create(nameof(this.ShapeRouteStyle), VASS.SrcConstants.ShapeLayoutShapeRouteStyle,
            this.ShapeRouteStyle);
        yield return this.Create(nameof(this.ShapeSplit), VASS.SrcConstants.ShapeLayoutShapeSplit, this.ShapeSplit);
        yield return this.Create(nameof(this.ShapeSplittable), VASS.SrcConstants.ShapeLayoutShapeSplittable,
            this.ShapeSplittable);
        yield return this.Create(nameof(this.ShapeDisplayLevel), VASS.SrcConstants.ShapeLayoutShapeDisplayLevel,
            this.ShapeDisplayLevel);
        yield return this.Create(nameof(this.Relationships), VASS.SrcConstants.ShapeLayoutRelationships,
            this.Relationships);
    }


    public static List<ShapeLayoutCells> GetCells(IVisio.Page page, IList<int> shapeids, VASS.CellValueType type)
    {
        var reader = ShapeLayoutCells_lazy_builder.Value;
        return reader.GetCellsSingleRow(page, shapeids, type);
    }

    public static ShapeLayoutCells GetCells(IVisio.Shape shape, VASS.CellValueType type)
    {
        var reader = ShapeLayoutCells_lazy_builder.Value;
        return reader.GetCellsSingleRow(shape, type);
    }

    private static readonly System.Lazy<ShapeLayoutCellsBuilder> ShapeLayoutCells_lazy_builder = new System.Lazy<ShapeLayoutCellsBuilder>();

    class ShapeLayoutCellsBuilder : VASS.CellGroups.CellGroupBuilder<ShapeLayoutCells>
    {

        public ShapeLayoutCellsBuilder() : base(VisioAutomation.ShapeSheet.CellGroups.CellGroupBuilderType.SingleRow)
        {
        }

        public override ShapeLayoutCells ToCellGroup(ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols)
        {
            var cells = new ShapeLayoutCells();
            var getcellvalue = VisioAutomation.ShapeSheet.CellGroups.CellGroup.row_to_cellgroup(row, cols);

            cells.ConnectorFixedCode = getcellvalue(nameof(ShapeLayoutCells.ConnectorFixedCode));
            cells.LineJumpCode = getcellvalue(nameof(ShapeLayoutCells.LineJumpCode));
            cells.LineJumpDirX = getcellvalue(nameof(ShapeLayoutCells.LineJumpDirX));
            cells.LineJumpDirY = getcellvalue(nameof(ShapeLayoutCells.LineJumpDirY));
            cells.LineJumpStyle = getcellvalue(nameof(ShapeLayoutCells.LineJumpStyle));
            cells.LineRouteExt = getcellvalue(nameof(ShapeLayoutCells.LineRouteExt));
            cells.ShapeFixedCode = getcellvalue(nameof(ShapeLayoutCells.ShapeFixedCode));
            cells.ShapePermeablePlace = getcellvalue(nameof(ShapeLayoutCells.ShapePermeablePlace));
            cells.ShapePermeableX = getcellvalue(nameof(ShapeLayoutCells.ShapePermeableX));
            cells.ShapePermeableY = getcellvalue(nameof(ShapeLayoutCells.ShapePermeableY));
            cells.ShapePlaceFlip = getcellvalue(nameof(ShapeLayoutCells.ShapePlaceFlip));
            cells.ShapePlaceStyle = getcellvalue(nameof(ShapeLayoutCells.ShapePlaceStyle));
            cells.ShapePlowCode = getcellvalue(nameof(ShapeLayoutCells.ShapePlowCode));
            cells.ShapeRouteStyle = getcellvalue(nameof(ShapeLayoutCells.ShapeRouteStyle));
            cells.ShapeSplit = getcellvalue(nameof(ShapeLayoutCells.ShapeSplit));
            cells.ShapeSplittable = getcellvalue(nameof(ShapeLayoutCells.ShapeSplittable));
            cells.ShapeDisplayLevel = getcellvalue(nameof(ShapeLayoutCells.ShapeDisplayLevel));
            cells.Relationships = getcellvalue(nameof(ShapeLayoutCells.Relationships));

            return cells;
        }
    }

}