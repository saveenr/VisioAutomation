using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class LayoutCells : VASS.CellGroups.CellGroup
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


        public static List<LayoutCells> GetCells(IVisio.Page page, IList<int> shapeids, VisioAutomation.Core.CellValueType type)
        {
            var reader = ShapeLayoutCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static LayoutCells GetCells(IVisio.Shape shape, VisioAutomation.Core.CellValueType type)
        {
            var reader = ShapeLayoutCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<ShapeLayoutCellsBuilder> ShapeLayoutCells_lazy_builder = new System.Lazy<ShapeLayoutCellsBuilder>();

        class ShapeLayoutCellsBuilder : VASS.CellGroups.CellGroupBuilder<LayoutCells>
        {

            public ShapeLayoutCellsBuilder() : base(VisioAutomation.ShapeSheet.CellGroups.CellGroupBuilderType.SingleRow)
            {
            }

            public override LayoutCells ToCellGroup(ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols)
            {
                var cells = new LayoutCells();
                var getcellvalue = VisioAutomation.ShapeSheet.CellGroups.CellGroup.row_to_cellgroup(row, cols);

                cells.ConnectorFixedCode = getcellvalue(nameof(LayoutCells.ConnectorFixedCode));
                cells.LineJumpCode = getcellvalue(nameof(LayoutCells.LineJumpCode));
                cells.LineJumpDirX = getcellvalue(nameof(LayoutCells.LineJumpDirX));
                cells.LineJumpDirY = getcellvalue(nameof(LayoutCells.LineJumpDirY));
                cells.LineJumpStyle = getcellvalue(nameof(LayoutCells.LineJumpStyle));
                cells.LineRouteExt = getcellvalue(nameof(LayoutCells.LineRouteExt));
                cells.ShapeFixedCode = getcellvalue(nameof(LayoutCells.ShapeFixedCode));
                cells.ShapePermeablePlace = getcellvalue(nameof(LayoutCells.ShapePermeablePlace));
                cells.ShapePermeableX = getcellvalue(nameof(LayoutCells.ShapePermeableX));
                cells.ShapePermeableY = getcellvalue(nameof(LayoutCells.ShapePermeableY));
                cells.ShapePlaceFlip = getcellvalue(nameof(LayoutCells.ShapePlaceFlip));
                cells.ShapePlaceStyle = getcellvalue(nameof(LayoutCells.ShapePlaceStyle));
                cells.ShapePlowCode = getcellvalue(nameof(LayoutCells.ShapePlowCode));
                cells.ShapeRouteStyle = getcellvalue(nameof(LayoutCells.ShapeRouteStyle));
                cells.ShapeSplit = getcellvalue(nameof(LayoutCells.ShapeSplit));
                cells.ShapeSplittable = getcellvalue(nameof(LayoutCells.ShapeSplittable));
                cells.ShapeDisplayLevel = getcellvalue(nameof(LayoutCells.ShapeDisplayLevel));
                cells.Relationships = getcellvalue(nameof(LayoutCells.Relationships));

                return cells;
            }
        }

    }
}