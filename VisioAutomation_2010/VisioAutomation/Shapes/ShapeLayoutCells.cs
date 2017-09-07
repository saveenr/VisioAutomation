using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio= Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class ShapeLayoutCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public ShapeSheet.CellData ConnectorFixedCode { get; set; }
        public ShapeSheet.CellData LineJumpCode { get; set; }
        public ShapeSheet.CellData LineJumpDirX { get; set; }
        public ShapeSheet.CellData LineJumpDirY { get; set; }
        public ShapeSheet.CellData LineJumpStyle { get; set; }
        public ShapeSheet.CellData LineRouteExt { get; set; }
        public ShapeSheet.CellData ShapeFixedCode { get; set; }
        public ShapeSheet.CellData ShapePermeablePlace { get; set; }
        public ShapeSheet.CellData ShapePermeableX { get; set; }
        public ShapeSheet.CellData ShapePermeableY { get; set; }
        public ShapeSheet.CellData ShapePlaceFlip { get; set; }
        public ShapeSheet.CellData ShapePlaceStyle { get; set; }
        public ShapeSheet.CellData ShapePlowCode { get; set; }
        public ShapeSheet.CellData ShapeRouteStyle { get; set; }
        public ShapeSheet.CellData ShapeSplit { get; set; }
        public ShapeSheet.CellData ShapeSplittable { get; set; }
        public ShapeSheet.CellData ShapeDisplayLevel { get; set; } // new in visio 2010
        public ShapeSheet.CellData Relationships { get; set; } // new in visio 2010

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