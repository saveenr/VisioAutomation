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
        public ShapeSheet.CellData FixedCode { get; set; }
        public ShapeSheet.CellData PermeablePlace { get; set; }
        public ShapeSheet.CellData PermeableX { get; set; }
        public ShapeSheet.CellData PermeableY { get; set; }
        public ShapeSheet.CellData PlaceFlip { get; set; }
        public ShapeSheet.CellData PlaceStyle { get; set; }
        public ShapeSheet.CellData PlowCode { get; set; }
        public ShapeSheet.CellData RouteStyle { get; set; }
        public ShapeSheet.CellData Split { get; set; }
        public ShapeSheet.CellData Splittable { get; set; }
        public ShapeSheet.CellData DisplayLevel { get; set; } // new in visio 2010
        public ShapeSheet.CellData Relationships { get; set; } // new in visio 2010

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutConFixedCode, this.ConFixedCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutConLineJumpCode, this.ConLineJumpCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutConLineJumpDirX, this.ConLineJumpDirX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutConLineJumpDirY, this.ConLineJumpDirY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutConLineJumpStyle, this.ConLineJumpStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutConLineRouteExt, this.ConLineRouteExt.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutFixedCode, this.FixedCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutPermeablePlace, this.PermeablePlace.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutPermeableX, this.PermeableX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutPermeableY, this.PermeableY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutPlaceFlip, this.PlaceFlip.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutPlaceStyle, this.PlaceStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutPlowCode, this.PlowCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutRouteStyle, this.RouteStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutSplit, this.Split.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutSplittable, this.Splittable.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutDisplayLevel, this.DisplayLevel.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShapeLayoutRelationships, this.Relationships.Formula);
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