using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.CellGroups.Queries;
using IVisio= Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.Layout
{
    public class ShapeLayoutCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public ShapeSheet.CellData ConFixedCode { get; set; }
        public ShapeSheet.CellData ConLineJumpCode { get; set; }
        public ShapeSheet.CellData ConLineJumpDirX { get; set; }
        public ShapeSheet.CellData ConLineJumpDirY { get; set; }
        public ShapeSheet.CellData ConLineJumpStyle { get; set; }
        public ShapeSheet.CellData ConLineRouteExt { get; set; }
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
        public ShapeSheet.CellData DisplayLevel { get; set; } // new in visio 2010
        public ShapeSheet.CellData Relationships { get; set; } // new in visio 2010

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


        public static IList<ShapeLayoutCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = ShapeLayoutCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static ShapeLayoutCells GetCells(IVisio.Shape shape)
        {
            var query = ShapeLayoutCells.lazy_query.Value;
            return query.GetCellGroup(shape);
        }

        private static System.Lazy<ShapeLayoutCellsQuery> lazy_query = new System.Lazy<ShapeLayoutCellsQuery>();



    }
}