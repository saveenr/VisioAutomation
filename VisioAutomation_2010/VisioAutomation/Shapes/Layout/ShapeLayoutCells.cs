using System.Collections.Generic;

namespace VisioAutomation.Shapes.Layout
{
    public class ShapeLayoutCells : ShapeSheet.CellGroups.CellGroup
    {
        public ShapeSheet.CellData<int> ConFixedCode { get; set; }
        public ShapeSheet.CellData<int> ConLineJumpCode { get; set; }
        public ShapeSheet.CellData<int> ConLineJumpDirX { get; set; }
        public ShapeSheet.CellData<int> ConLineJumpDirY { get; set; }
        public ShapeSheet.CellData<int> ConLineJumpStyle { get; set; }
        public ShapeSheet.CellData<int> ConLineRouteExt { get; set; }
        public ShapeSheet.CellData<int> ShapeFixedCode { get; set; }
        public ShapeSheet.CellData<int> ShapePermeablePlace { get; set; }
        public ShapeSheet.CellData<int> ShapePermeableX { get; set; }
        public ShapeSheet.CellData<int> ShapePermeableY { get; set; }
        public ShapeSheet.CellData<int> ShapePlaceFlip { get; set; }
        public ShapeSheet.CellData<int> ShapePlaceStyle { get; set; }
        public ShapeSheet.CellData<int> ShapePlowCode { get; set; }
        public ShapeSheet.CellData<int> ShapeRouteStyle { get; set; }
        public ShapeSheet.CellData<int> ShapeSplit { get; set; }
        public ShapeSheet.CellData<int> ShapeSplittable { get; set; }
        public ShapeSheet.CellData<int> DisplayLevel { get; set; } // new in visio 2010
        public ShapeSheet.CellData<int> Relationships { get; set; } // new in visio 2010

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.ConFixedCode, this.ConFixedCode.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ConLineJumpCode, this.ConLineJumpCode.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ConLineJumpDirX, this.ConLineJumpDirX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ConLineJumpDirY, this.ConLineJumpDirY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ConLineJumpStyle, this.ConLineJumpStyle.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ConLineRouteExt, this.ConLineRouteExt.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapeFixedCode, this.ShapeFixedCode.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapePermeablePlace, this.ShapePermeablePlace.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapePermeableX, this.ShapePermeableX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapePermeableY, this.ShapePermeableY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapePlaceFlip, this.ShapePlaceFlip.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapePlaceStyle, this.ShapePlaceStyle.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapePlowCode, this.ShapePlowCode.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapeRouteStyle, this.ShapeRouteStyle.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapeSplit, this.ShapeSplit.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShapeSplittable, this.ShapeSplittable.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.DisplayLevel, this.DisplayLevel.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Relationships, this.Relationships.Formula);
            }
        }


        public static IList<ShapeLayoutCells> GetCells(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var query = ShapeLayoutCells.lazy_query.Value;
            return ShapeSheet.CellGroups.CellGroup._GetCells<ShapeLayoutCells, double>(page, shapeids, query, query.GetCells);
        }

        public static ShapeLayoutCells GetCells(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = ShapeLayoutCells.lazy_query.Value;
            return ShapeSheet.CellGroups.CellGroup._GetCells<ShapeLayoutCells, double>(shape, query, query.GetCells);
        }

        private static System.Lazy<ShapeSheet.Query.Common.ShapeLayoutCellsQuery> lazy_query = new System.Lazy<ShapeSheet.Query.Common.ShapeLayoutCellsQuery>();



    }
}