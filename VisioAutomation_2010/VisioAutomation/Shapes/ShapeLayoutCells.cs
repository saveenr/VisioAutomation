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

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutConnectorFixedCode, this.ConnectorFixedCode.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutLineJumpCode, this.LineJumpCode.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutLineJumpDirX, this.LineJumpDirX.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutLineJumpDirY, this.LineJumpDirY.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutLineJumpStyle, this.LineJumpStyle.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutLineRouteExt, this.LineRouteExt.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapeFixedCode, this.ShapeFixedCode.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapePermeablePlace, this.ShapePermeablePlace.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapePermeableX, this.ShapePermeableX.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapePermeableY, this.ShapePermeableY.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapePlaceFlip, this.ShapePlaceFlip.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapePlaceStyle, this.ShapePlaceStyle.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapePlowCode, this.ShapePlowCode.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapeRouteStyle, this.ShapeRouteStyle.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapeSplit, this.ShapeSplit.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapeSplittable, this.ShapeSplittable.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapeDisplayLevel, this.ShapeDisplayLevel.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutRelationships, this.Relationships.Value);
            }
        }


        public static List<ShapeLayoutCells> GetCells(IVisio.Page page, IList<int> shapeids, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = ShapeLayoutCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids, cvt);
        }

        public static ShapeLayoutCells GetCells(IVisio.Shape shape, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = ShapeLayoutCells.lazy_query.Value;
            return query.GetCellGroup(shape, cvt);
        }

        private static readonly System.Lazy<ShapeLayoutCellsReader> lazy_query = new System.Lazy<ShapeLayoutCellsReader>();
    }
}