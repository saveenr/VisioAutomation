using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
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

        public override IEnumerable<SRCFormulaPair> SRCFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.ConFixedCode, this.ConFixedCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ConLineJumpCode, this.ConLineJumpCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ConLineJumpDirX, this.ConLineJumpDirX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ConLineJumpDirY, this.ConLineJumpDirY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ConLineJumpStyle, this.ConLineJumpStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ConLineRouteExt, this.ConLineRouteExt.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeFixedCode, this.ShapeFixedCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapePermeablePlace, this.ShapePermeablePlace.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapePermeableX, this.ShapePermeableX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapePermeableY, this.ShapePermeableY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapePlaceFlip, this.ShapePlaceFlip.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapePlaceStyle, this.ShapePlaceStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapePlowCode, this.ShapePlowCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeRouteStyle, this.ShapeRouteStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeSplit, this.ShapeSplit.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeSplittable, this.ShapeSplittable.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.DisplayLevel, this.DisplayLevel.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.Relationships, this.Relationships.Formula);
            }
        }


        public static List<ShapeLayoutCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = ShapeLayoutCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static ShapeLayoutCells GetCells(IVisio.Shape shape)
        {
            var query = ShapeLayoutCells.lazy_query.Value;
            return query.GetCellGroup(shape);
        }

        private static readonly System.Lazy<ShapeLayoutCellsReader> lazy_query = new System.Lazy<ShapeLayoutCellsReader>();
    }
}