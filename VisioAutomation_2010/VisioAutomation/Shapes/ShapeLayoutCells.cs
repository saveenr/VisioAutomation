using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class ShapeLayoutCells : VASS.CellGroups.CellGroup
    {
        public VisioAutomation.Core.CellValue ConnectorFixedCode { get; set; }
        public VisioAutomation.Core.CellValue LineJumpCode { get; set; }
        public VisioAutomation.Core.CellValue LineJumpDirX { get; set; }
        public VisioAutomation.Core.CellValue LineJumpDirY { get; set; }
        public VisioAutomation.Core.CellValue LineJumpStyle { get; set; }
        public VisioAutomation.Core.CellValue LineRouteExt { get; set; }
        public VisioAutomation.Core.CellValue ShapeFixedCode { get; set; }
        public VisioAutomation.Core.CellValue ShapePermeablePlace { get; set; }
        public VisioAutomation.Core.CellValue ShapePermeableX { get; set; }
        public VisioAutomation.Core.CellValue ShapePermeableY { get; set; }
        public VisioAutomation.Core.CellValue ShapePlaceFlip { get; set; }
        public VisioAutomation.Core.CellValue ShapePlaceStyle { get; set; }
        public VisioAutomation.Core.CellValue ShapePlowCode { get; set; }
        public VisioAutomation.Core.CellValue ShapeRouteStyle { get; set; }
        public VisioAutomation.Core.CellValue ShapeSplit { get; set; }
        public VisioAutomation.Core.CellValue ShapeSplittable { get; set; }
        public VisioAutomation.Core.CellValue ShapeDisplayLevel { get; set; } // new in visio 2010
        public VisioAutomation.Core.CellValue Relationships { get; set; } // new in visio 2010

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.ConnectorFixedCode), VisioAutomation.Core.SrcConstants.ShapeLayoutConnectorFixedCode,
                this.ConnectorFixedCode);
            yield return this.Create(nameof(this.LineJumpCode), VisioAutomation.Core.SrcConstants.ShapeLayoutLineJumpCode, this.LineJumpCode);
            yield return this.Create(nameof(this.LineJumpDirX), VisioAutomation.Core.SrcConstants.ShapeLayoutLineJumpDirX, this.LineJumpDirX);
            yield return this.Create(nameof(this.LineJumpDirY), VisioAutomation.Core.SrcConstants.ShapeLayoutLineJumpDirY, this.LineJumpDirY);
            yield return this.Create(nameof(this.LineJumpStyle), VisioAutomation.Core.SrcConstants.ShapeLayoutLineJumpStyle,
                this.LineJumpStyle);
            yield return this.Create(nameof(this.LineRouteExt), VisioAutomation.Core.SrcConstants.ShapeLayoutLineRouteExt, this.LineRouteExt);
            yield return this.Create(nameof(this.ShapeFixedCode), VisioAutomation.Core.SrcConstants.ShapeLayoutShapeFixedCode,
                this.ShapeFixedCode);
            yield return this.Create(nameof(this.ShapePermeablePlace), VisioAutomation.Core.SrcConstants.ShapeLayoutShapePermeablePlace,
                this.ShapePermeablePlace);
            yield return this.Create(nameof(this.ShapePermeableX), VisioAutomation.Core.SrcConstants.ShapeLayoutShapePermeableX,
                this.ShapePermeableX);
            yield return this.Create(nameof(this.ShapePermeableY), VisioAutomation.Core.SrcConstants.ShapeLayoutShapePermeableY,
                this.ShapePermeableY);
            yield return this.Create(nameof(this.ShapePlaceFlip), VisioAutomation.Core.SrcConstants.ShapeLayoutShapePlaceFlip,
                this.ShapePlaceFlip);
            yield return this.Create(nameof(this.ShapePlaceStyle), VisioAutomation.Core.SrcConstants.ShapeLayoutShapePlaceStyle,
                this.ShapePlaceStyle);
            yield return this.Create(nameof(this.ShapePlowCode), VisioAutomation.Core.SrcConstants.ShapeLayoutShapePlowCode,
                this.ShapePlowCode);
            yield return this.Create(nameof(this.ShapeRouteStyle), VisioAutomation.Core.SrcConstants.ShapeLayoutShapeRouteStyle,
                this.ShapeRouteStyle);
            yield return this.Create(nameof(this.ShapeSplit), VisioAutomation.Core.SrcConstants.ShapeLayoutShapeSplit, this.ShapeSplit);
            yield return this.Create(nameof(this.ShapeSplittable), VisioAutomation.Core.SrcConstants.ShapeLayoutShapeSplittable,
                this.ShapeSplittable);
            yield return this.Create(nameof(this.ShapeDisplayLevel), VisioAutomation.Core.SrcConstants.ShapeLayoutShapeDisplayLevel,
                this.ShapeDisplayLevel);
            yield return this.Create(nameof(this.Relationships), VisioAutomation.Core.SrcConstants.ShapeLayoutRelationships,
                this.Relationships);
        }


        public static List<ShapeLayoutCells> GetCells(IVisio.Page page, IList<int> shapeids, VisioAutomation.Core.CellValueType type)
        {
            var reader = ShapeLayoutCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static ShapeLayoutCells GetCells(IVisio.Shape shape, VisioAutomation.Core.CellValueType type)
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
}