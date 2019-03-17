using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class ShapeLayoutCells : CellGroup
    {
        public VASS.CellValueLiteral ConnectorFixedCode { get; set; }
        public VASS.CellValueLiteral LineJumpCode { get; set; }
        public VASS.CellValueLiteral LineJumpDirX { get; set; }
        public VASS.CellValueLiteral LineJumpDirY { get; set; }
        public VASS.CellValueLiteral LineJumpStyle { get; set; }
        public VASS.CellValueLiteral LineRouteExt { get; set; }
        public VASS.CellValueLiteral ShapeFixedCode { get; set; }
        public VASS.CellValueLiteral ShapePermeablePlace { get; set; }
        public VASS.CellValueLiteral ShapePermeableX { get; set; }
        public VASS.CellValueLiteral ShapePermeableY { get; set; }
        public VASS.CellValueLiteral ShapePlaceFlip { get; set; }
        public VASS.CellValueLiteral ShapePlaceStyle { get; set; }
        public VASS.CellValueLiteral ShapePlowCode { get; set; }
        public VASS.CellValueLiteral ShapeRouteStyle { get; set; }
        public VASS.CellValueLiteral ShapeSplit { get; set; }
        public VASS.CellValueLiteral ShapeSplittable { get; set; }
        public VASS.CellValueLiteral ShapeDisplayLevel { get; set; } // new in visio 2010
        public VASS.CellValueLiteral Relationships { get; set; } // new in visio 2010

        public override IEnumerable<CellMetadataItem> CellMetadata
        {
            get
            {


                yield return this.Create(nameof(this.ConnectorFixedCode), VASS.SrcConstants.ShapeLayoutConnectorFixedCode, this.ConnectorFixedCode);
                yield return this.Create(nameof(this.LineJumpCode), VASS.SrcConstants.ShapeLayoutLineJumpCode, this.LineJumpCode);
                yield return this.Create(nameof(this.LineJumpDirX), VASS.SrcConstants.ShapeLayoutLineJumpDirX, this.LineJumpDirX);
                yield return this.Create(nameof(this.LineJumpDirY), VASS.SrcConstants.ShapeLayoutLineJumpDirY, this.LineJumpDirY);
                yield return this.Create(nameof(this.LineJumpStyle), VASS.SrcConstants.ShapeLayoutLineJumpStyle, this.LineJumpStyle);
                yield return this.Create(nameof(this.LineRouteExt), VASS.SrcConstants.ShapeLayoutLineRouteExt, this.LineRouteExt);
                yield return this.Create(nameof(this.ShapeFixedCode), VASS.SrcConstants.ShapeLayoutShapeFixedCode, this.ShapeFixedCode);
                yield return this.Create(nameof(this.ShapePermeablePlace), VASS.SrcConstants.ShapeLayoutShapePermeablePlace, this.ShapePermeablePlace);
                yield return this.Create(nameof(this.ShapePermeableX), VASS.SrcConstants.ShapeLayoutShapePermeableX, this.ShapePermeableX);
                yield return this.Create(nameof(this.ShapePermeableY), VASS.SrcConstants.ShapeLayoutShapePermeableY, this.ShapePermeableY);
                yield return this.Create(nameof(this.ShapePlaceFlip), VASS.SrcConstants.ShapeLayoutShapePlaceFlip, this.ShapePlaceFlip);
                yield return this.Create(nameof(this.ShapePlaceStyle), VASS.SrcConstants.ShapeLayoutShapePlaceStyle, this.ShapePlaceStyle);
                yield return this.Create(nameof(this.ShapePlowCode), VASS.SrcConstants.ShapeLayoutShapePlowCode, this.ShapePlowCode);
                yield return this.Create(nameof(this.ShapeRouteStyle), VASS.SrcConstants.ShapeLayoutShapeRouteStyle, this.ShapeRouteStyle);
                yield return this.Create(nameof(this.ShapeSplit), VASS.SrcConstants.ShapeLayoutShapeSplit, this.ShapeSplit);
                yield return this.Create(nameof(this.ShapeSplittable), VASS.SrcConstants.ShapeLayoutShapeSplittable, this.ShapeSplittable);
                yield return this.Create(nameof(this.ShapeDisplayLevel), VASS.SrcConstants.ShapeLayoutShapeDisplayLevel, this.ShapeDisplayLevel);
                yield return this.Create(nameof(this.Relationships), VASS.SrcConstants.ShapeLayoutRelationships, this.Relationships);
            }
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

        class ShapeLayoutCellsBuilder : CellGroupBuilder<ShapeLayoutCells>
        {

            public ShapeLayoutCellsBuilder() : base(VisioAutomation.ShapeSheet.CellGroups.CellGroupBuilderType.SingleRow)
            {
            }

            public override ShapeLayoutCells ToCellGroup(ShapeSheet.Query.ShapeCellsRow<string> row, VisioAutomation.ShapeSheet.Query.ColumnList cols)
            {
                var cells = new ShapeLayoutCells();
                var getcellvalue = VisioAutomation.ShapeSheet.CellGroups.CellGroup.gcf(row, cols);

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
}