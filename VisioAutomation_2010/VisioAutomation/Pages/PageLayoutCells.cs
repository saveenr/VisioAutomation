using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;

namespace VisioAutomation.Pages
{
    public class PageLayoutCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public ShapeSheet.CellData PageLayoutAvenueSizeX { get; set; }
        public ShapeSheet.CellData PageLayoutAvenueSizeY { get; set; }
        public ShapeSheet.CellData PageLayoutBlockSizeX { get; set; }
        public ShapeSheet.CellData PageLayoutBlockSizeY { get; set; }
        public ShapeSheet.CellData PageLayoutCtrlAsInput { get; set; }
        public ShapeSheet.CellData PageLayoutDynamicsOff { get; set; }
        public ShapeSheet.CellData PageLayoutEnableGrid { get; set; }
        public ShapeSheet.CellData PageLayoutLineAdjustFrom { get; set; }
        public ShapeSheet.CellData PageLayoutLineAdjustTo { get; set; }
        public ShapeSheet.CellData PageLayoutLineJumpCode { get; set; }
        public ShapeSheet.CellData PageLayoutLineJumpFactorX { get; set; }
        public ShapeSheet.CellData PageLayoutLineJumpFactorY { get; set; }
        public ShapeSheet.CellData PageLayoutLineJumpStyle { get; set; }
        public ShapeSheet.CellData PageLayoutLineRouteExt { get; set; }
        public ShapeSheet.CellData PageLayoutLineToLineX { get; set; }
        public ShapeSheet.CellData PageLayoutLineToLineY { get; set; }
        public ShapeSheet.CellData PageLayoutLineToNodeX { get; set; }
        public ShapeSheet.CellData PageLayoutLineToNodeY { get; set; }
        public ShapeSheet.CellData PageLayoutLineJumpDirX { get; set; }
        public ShapeSheet.CellData PageLayoutLineJumpDirY { get; set; }
        public ShapeSheet.CellData PageLayoutPageShapeSplit { get; set; }
        public ShapeSheet.CellData PageLayoutPlaceDepth { get; set; }
        public ShapeSheet.CellData PageLayoutPlaceFlip { get; set; }
        public ShapeSheet.CellData PageLayoutPlaceStyle { get; set; }
        public ShapeSheet.CellData PageLayoutPlowCode { get; set; }
        public ShapeSheet.CellData PageLayoutResizePage { get; set; }
        public ShapeSheet.CellData PageLayoutRouteStyle { get; set; }
        public ShapeSheet.CellData PageLayoutAvoidPageBreaks { get; set; } // new in visio 2010

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutAvenueSizeX, this.PageLayoutAvenueSizeX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutAvenueSizeY, this.PageLayoutAvenueSizeY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutBlockSizeX, this.PageLayoutBlockSizeX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutBlockSizeY, this.PageLayoutBlockSizeY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutControlAsInput, this.PageLayoutCtrlAsInput.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutDynamicsOff, this.PageLayoutDynamicsOff.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutEnableGrid, this.PageLayoutEnableGrid.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineAdjustFrom, this.PageLayoutLineAdjustFrom.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineAdjustTo, this.PageLayoutLineAdjustTo.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpCode, this.PageLayoutLineJumpCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpFactorX, this.PageLayoutLineJumpFactorX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpFactorY, this.PageLayoutLineJumpFactorY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpStyle, this.PageLayoutLineJumpStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineRouteExt, this.PageLayoutLineRouteExt.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineToLineX, this.PageLayoutLineToLineX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineToLineY, this.PageLayoutLineToLineY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineToNodeX, this.PageLayoutLineToNodeX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineToNodeY, this.PageLayoutLineToNodeY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpDirX, this.PageLayoutLineJumpDirX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpDirY, this.PageLayoutLineJumpDirY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutShapeSplit, this.PageLayoutPageShapeSplit.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlaceDepth, this.PageLayoutPlaceDepth.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlaceFlip, this.PageLayoutPlaceFlip.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlaceStyle, this.PageLayoutPlaceStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlowCode, this.PageLayoutPlowCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutResizePage, this.PageLayoutResizePage.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutRouteStyle, this.PageLayoutRouteStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutAvoidPageBreaks, this.PageLayoutAvoidPageBreaks.Formula);
            }
        }

        public static PageLayoutCells GetCells(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = PageLayoutCells.lazy_query.Value;
            return query.GetCellGroup(shape);
        }

        private static readonly System.Lazy<PageLayoutCellsReader> lazy_query = new System.Lazy<PageLayoutCellsReader>();
    }
}