using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio= Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

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

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutConnectorFixedCode, this.ConnectorFixedCode);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutLineJumpCode, this.LineJumpCode);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutLineJumpDirX, this.LineJumpDirX);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutLineJumpDirY, this.LineJumpDirY);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutLineJumpStyle, this.LineJumpStyle);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutLineRouteExt, this.LineRouteExt);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapeFixedCode, this.ShapeFixedCode);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapePermeablePlace, this.ShapePermeablePlace);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapePermeableX, this.ShapePermeableX);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapePermeableY, this.ShapePermeableY);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapePlaceFlip, this.ShapePlaceFlip);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapePlaceStyle, this.ShapePlaceStyle);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapePlowCode, this.ShapePlowCode);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapeRouteStyle, this.ShapeRouteStyle);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapeSplit, this.ShapeSplit);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapeSplittable, this.ShapeSplittable);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapeDisplayLevel, this.ShapeDisplayLevel);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutRelationships, this.Relationships);
            }
        }
    }
}