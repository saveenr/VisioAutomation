using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio= Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class ShapeLayoutCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral ConnectorFixedCode { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpCode { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpDirX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpDirY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpStyle { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineRouteExt { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapeFixedCode { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapePermeablePlace { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapePermeableX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapePermeableY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapePlaceFlip { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapePlaceStyle { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapePlowCode { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapeRouteStyle { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapeSplit { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapeSplittable { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapeDisplayLevel { get; set; } // new in visio 2010
        public VisioAutomation.ShapeSheet.CellValueLiteral Relationships { get; set; } // new in visio 2010

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutConnectorFixedCode, this.ConnectorFixedCode.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutLineJumpCode, this.LineJumpCode.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutLineJumpDirX, this.LineJumpDirX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutLineJumpDirY, this.LineJumpDirY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutLineJumpStyle, this.LineJumpStyle.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutLineRouteExt, this.LineRouteExt.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapeFixedCode, this.ShapeFixedCode.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapePermeablePlace, this.ShapePermeablePlace.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapePermeableX, this.ShapePermeableX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapePermeableY, this.ShapePermeableY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapePlaceFlip, this.ShapePlaceFlip.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapePlaceStyle, this.ShapePlaceStyle.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapePlowCode, this.ShapePlowCode.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapeRouteStyle, this.ShapeRouteStyle.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapeSplit, this.ShapeSplit.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapeSplittable, this.ShapeSplittable.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapeDisplayLevel, this.ShapeDisplayLevel.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutRelationships, this.Relationships.Value);
            }
        }
        
        public static List<ShapeLayoutCells> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var query = ShapeLayoutCells.lazy_query.Value;
            return query.GetFormulas(page, shapeids);
        }

        public static List<ShapeLayoutCells> GetResults(IVisio.Page page, IList<int> shapeids)
        {
            var query = ShapeLayoutCells.lazy_query.Value;
            return query.GetResults(page, shapeids);
        }

        public static ShapeLayoutCells GetFormulas(IVisio.Shape shape)
        {
            var query = ShapeLayoutCells.lazy_query.Value;
            return query.GetFormulas(shape);
        }

        public static ShapeLayoutCells GetResults(IVisio.Shape shape)
        {
            var query = ShapeLayoutCells.lazy_query.Value;
            return query.GetResults(shape);
        }

        private static readonly System.Lazy<ShapeLayoutCellsReader> lazy_query = new System.Lazy<ShapeLayoutCellsReader>();
    }
}