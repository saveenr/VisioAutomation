using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class ShapeLayoutCells : CellGroup
    {
        public CellValueLiteral ConnectorFixedCode { get; set; }
        public CellValueLiteral LineJumpCode { get; set; }
        public CellValueLiteral LineJumpDirX { get; set; }
        public CellValueLiteral LineJumpDirY { get; set; }
        public CellValueLiteral LineJumpStyle { get; set; }
        public CellValueLiteral LineRouteExt { get; set; }
        public CellValueLiteral ShapeFixedCode { get; set; }
        public CellValueLiteral ShapePermeablePlace { get; set; }
        public CellValueLiteral ShapePermeableX { get; set; }
        public CellValueLiteral ShapePermeableY { get; set; }
        public CellValueLiteral ShapePlaceFlip { get; set; }
        public CellValueLiteral ShapePlaceStyle { get; set; }
        public CellValueLiteral ShapePlowCode { get; set; }
        public CellValueLiteral ShapeRouteStyle { get; set; }
        public CellValueLiteral ShapeSplit { get; set; }
        public CellValueLiteral ShapeSplittable { get; set; }
        public CellValueLiteral ShapeDisplayLevel { get; set; } // new in visio 2010
        public CellValueLiteral Relationships { get; set; } // new in visio 2010

        public override IEnumerable<CellMetadataItem> CellMetadata
        {
            get
            {


                yield return CellMetadataItem.Create(nameof(this.ConnectorFixedCode), SrcConstants.ShapeLayoutConnectorFixedCode, this.ConnectorFixedCode);
                yield return CellMetadataItem.Create(nameof(this.LineJumpCode), SrcConstants.ShapeLayoutLineJumpCode, this.LineJumpCode);
                yield return CellMetadataItem.Create(nameof(this.LineJumpDirX), SrcConstants.ShapeLayoutLineJumpDirX, this.LineJumpDirX);
                yield return CellMetadataItem.Create(nameof(this.LineJumpDirY), SrcConstants.ShapeLayoutLineJumpDirY, this.LineJumpDirY);
                yield return CellMetadataItem.Create(nameof(this.LineJumpStyle), SrcConstants.ShapeLayoutLineJumpStyle, this.LineJumpStyle);
                yield return CellMetadataItem.Create(nameof(this.LineRouteExt), SrcConstants.ShapeLayoutLineRouteExt, this.LineRouteExt);
                yield return CellMetadataItem.Create(nameof(this.ShapeFixedCode), SrcConstants.ShapeLayoutShapeFixedCode, this.ShapeFixedCode);
                yield return CellMetadataItem.Create(nameof(this.ShapePermeablePlace), SrcConstants.ShapeLayoutShapePermeablePlace, this.ShapePermeablePlace);
                yield return CellMetadataItem.Create(nameof(this.ShapePermeableX), SrcConstants.ShapeLayoutShapePermeableX, this.ShapePermeableX);
                yield return CellMetadataItem.Create(nameof(this.ShapePermeableY), SrcConstants.ShapeLayoutShapePermeableY, this.ShapePermeableY);
                yield return CellMetadataItem.Create(nameof(this.ShapePlaceFlip), SrcConstants.ShapeLayoutShapePlaceFlip, this.ShapePlaceFlip);
                yield return CellMetadataItem.Create(nameof(this.ShapePlaceStyle), SrcConstants.ShapeLayoutShapePlaceStyle, this.ShapePlaceStyle);
                yield return CellMetadataItem.Create(nameof(this.ShapePlowCode), SrcConstants.ShapeLayoutShapePlowCode, this.ShapePlowCode);
                yield return CellMetadataItem.Create(nameof(this.ShapeRouteStyle), SrcConstants.ShapeLayoutShapeRouteStyle, this.ShapeRouteStyle);
                yield return CellMetadataItem.Create(nameof(this.ShapeSplit), SrcConstants.ShapeLayoutShapeSplit, this.ShapeSplit);
                yield return CellMetadataItem.Create(nameof(this.ShapeSplittable), SrcConstants.ShapeLayoutShapeSplittable, this.ShapeSplittable);
                yield return CellMetadataItem.Create(nameof(this.ShapeDisplayLevel), SrcConstants.ShapeLayoutShapeDisplayLevel, this.ShapeDisplayLevel);
                yield return CellMetadataItem.Create(nameof(this.Relationships), SrcConstants.ShapeLayoutRelationships, this.Relationships);
            }
        }


        public static List<ShapeLayoutCells> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = ShapeLayoutCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static ShapeLayoutCells GetCells(IVisio.Shape shape, CellValueType type)
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

            public override ShapeLayoutCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row, VisioAutomation.ShapeSheet.Query.ColumnList cols)
            {
                var cells = new ShapeLayoutCells();

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }

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