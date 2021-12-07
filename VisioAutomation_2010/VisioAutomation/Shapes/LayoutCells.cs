using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellRecords;
using VACG = VisioAutomation.ShapeSheet.CellGroups;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class LayoutCells : VACG.CellRecord
    {
        public Core.CellValue ConnectorFixedCode { get; set; }
        public Core.CellValue LineJumpCode { get; set; }
        public Core.CellValue LineJumpDirX { get; set; }
        public Core.CellValue LineJumpDirY { get; set; }
        public Core.CellValue LineJumpStyle { get; set; }
        public Core.CellValue LineRouteExt { get; set; }
        public Core.CellValue ShapeFixedCode { get; set; }
        public Core.CellValue ShapePermeablePlace { get; set; }
        public Core.CellValue ShapePermeableX { get; set; }
        public Core.CellValue ShapePermeableY { get; set; }
        public Core.CellValue ShapePlaceFlip { get; set; }
        public Core.CellValue ShapePlaceStyle { get; set; }
        public Core.CellValue ShapePlowCode { get; set; }
        public Core.CellValue ShapeRouteStyle { get; set; }
        public Core.CellValue ShapeSplit { get; set; }
        public Core.CellValue ShapeSplittable { get; set; }
        public Core.CellValue ShapeDisplayLevel { get; set; } // new in visio 2010
        public Core.CellValue Relationships { get; set; } // new in visio 2010

        public override IEnumerable<VACG.CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.ConnectorFixedCode), Core.SrcConstants.ShapeLayoutConnectorFixedCode,
                this.ConnectorFixedCode);
            yield return this._create(nameof(this.LineJumpCode), Core.SrcConstants.ShapeLayoutLineJumpCode,
                this.LineJumpCode);
            yield return this._create(nameof(this.LineJumpDirX), Core.SrcConstants.ShapeLayoutLineJumpDirX,
                this.LineJumpDirX);
            yield return this._create(nameof(this.LineJumpDirY), Core.SrcConstants.ShapeLayoutLineJumpDirY,
                this.LineJumpDirY);
            yield return this._create(nameof(this.LineJumpStyle), Core.SrcConstants.ShapeLayoutLineJumpStyle,
                this.LineJumpStyle);
            yield return this._create(nameof(this.LineRouteExt), Core.SrcConstants.ShapeLayoutLineRouteExt,
                this.LineRouteExt);
            yield return this._create(nameof(this.ShapeFixedCode), Core.SrcConstants.ShapeLayoutShapeFixedCode,
                this.ShapeFixedCode);
            yield return this._create(nameof(this.ShapePermeablePlace),
                Core.SrcConstants.ShapeLayoutShapePermeablePlace,
                this.ShapePermeablePlace);
            yield return this._create(nameof(this.ShapePermeableX), Core.SrcConstants.ShapeLayoutShapePermeableX,
                this.ShapePermeableX);
            yield return this._create(nameof(this.ShapePermeableY), Core.SrcConstants.ShapeLayoutShapePermeableY,
                this.ShapePermeableY);
            yield return this._create(nameof(this.ShapePlaceFlip), Core.SrcConstants.ShapeLayoutShapePlaceFlip,
                this.ShapePlaceFlip);
            yield return this._create(nameof(this.ShapePlaceStyle), Core.SrcConstants.ShapeLayoutShapePlaceStyle,
                this.ShapePlaceStyle);
            yield return this._create(nameof(this.ShapePlowCode), Core.SrcConstants.ShapeLayoutShapePlowCode,
                this.ShapePlowCode);
            yield return this._create(nameof(this.ShapeRouteStyle), Core.SrcConstants.ShapeLayoutShapeRouteStyle,
                this.ShapeRouteStyle);
            yield return this._create(nameof(this.ShapeSplit), Core.SrcConstants.ShapeLayoutShapeSplit,
                this.ShapeSplit);
            yield return this._create(nameof(this.ShapeSplittable), Core.SrcConstants.ShapeLayoutShapeSplittable,
                this.ShapeSplittable);
            yield return this._create(nameof(this.ShapeDisplayLevel), Core.SrcConstants.ShapeLayoutShapeDisplayLevel,
                this.ShapeDisplayLevel);
            yield return this._create(nameof(this.Relationships), Core.SrcConstants.ShapeLayoutRelationships,
                this.Relationships);
        }


        public static List<LayoutCells> GetCells(IVisio.Page page, IList<int> shapeids, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsMultipleShapesSingleRow(page, shapeids, type);
        }

        public static LayoutCells GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleShapeSingleRow(shape, type);
        }

        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();

        class Builder : CellGroupBuilder<LayoutCells>
        {
            public Builder() : base(VACG.CellGroupBuilderType.SingleRow)
            {
            }

            public override LayoutCells ToCellGroup(VASS.Data.DataRow<string> row, VASS.Data.DataColumnCollection cols)
            {
                var cells = new LayoutCells();
                var getcellvalue = queryrow_to_cellgroup(row, cols);

                cells.ConnectorFixedCode = getcellvalue(nameof(ConnectorFixedCode));
                cells.LineJumpCode = getcellvalue(nameof(LineJumpCode));
                cells.LineJumpDirX = getcellvalue(nameof(LineJumpDirX));
                cells.LineJumpDirY = getcellvalue(nameof(LineJumpDirY));
                cells.LineJumpStyle = getcellvalue(nameof(LineJumpStyle));
                cells.LineRouteExt = getcellvalue(nameof(LineRouteExt));
                cells.ShapeFixedCode = getcellvalue(nameof(ShapeFixedCode));
                cells.ShapePermeablePlace = getcellvalue(nameof(ShapePermeablePlace));
                cells.ShapePermeableX = getcellvalue(nameof(ShapePermeableX));
                cells.ShapePermeableY = getcellvalue(nameof(ShapePermeableY));
                cells.ShapePlaceFlip = getcellvalue(nameof(ShapePlaceFlip));
                cells.ShapePlaceStyle = getcellvalue(nameof(ShapePlaceStyle));
                cells.ShapePlowCode = getcellvalue(nameof(ShapePlowCode));
                cells.ShapeRouteStyle = getcellvalue(nameof(ShapeRouteStyle));
                cells.ShapeSplit = getcellvalue(nameof(ShapeSplit));
                cells.ShapeSplittable = getcellvalue(nameof(ShapeSplittable));
                cells.ShapeDisplayLevel = getcellvalue(nameof(ShapeDisplayLevel));
                cells.Relationships = getcellvalue(nameof(Relationships));

                return cells;
            }
        }
    }
}