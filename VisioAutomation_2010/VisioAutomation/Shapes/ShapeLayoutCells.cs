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
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutConnectorFixedCode, this.ConnectorFixedCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutLineJumpCode, this.LineJumpCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutLineJumpDirX, this.LineJumpDirX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutLineJumpDirY, this.LineJumpDirY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutLineJumpStyle, this.LineJumpStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutLineRouteExt, this.LineRouteExt.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapeFixedCode, this.ShapeFixedCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapePermeablePlace, this.ShapePermeablePlace.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapePermeableX, this.ShapePermeableX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapePermeableY, this.ShapePermeableY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapePlaceFlip, this.ShapePlaceFlip.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapePlaceStyle, this.ShapePlaceStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapePlowCode, this.ShapePlowCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapeRouteStyle, this.ShapeRouteStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapeSplit, this.ShapeSplit.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapeSplittable, this.ShapeSplittable.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutShapeDisplayLevel, this.ShapeDisplayLevel.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutRelationships, this.Relationships.Formula);
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