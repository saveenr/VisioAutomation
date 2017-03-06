using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PageCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public ShapeSheet.CellData PrintLeftMargin { get; set; }
        public ShapeSheet.CellData PrintCenterX { get; set; }
        public ShapeSheet.CellData PrintCenterY { get; set; }
        public ShapeSheet.CellData PrintOnPage { get; set; }
        public ShapeSheet.CellData PrintBottomMargin { get; set; }
        public ShapeSheet.CellData PrintRightMargin { get; set; }
        public ShapeSheet.CellData PrintPagesX { get; set; }
        public ShapeSheet.CellData PrintPagesY { get; set; }
        public ShapeSheet.CellData PrintTopMargin { get; set; }
        public ShapeSheet.CellData PrintPaperKind { get; set; }
        public ShapeSheet.CellData PrintGrid { get; set; }
        public ShapeSheet.CellData PrintPageOrientation { get; set; }
        public ShapeSheet.CellData PrintScaleX { get; set; }
        public ShapeSheet.CellData PrintScaleY { get; set; }
        public ShapeSheet.CellData PrintPaperSource { get; set; }
        public ShapeSheet.CellData PageDrawingScale { get; set; }
        public ShapeSheet.CellData PageDrawingScaleType { get; set; }
        public ShapeSheet.CellData PageDrawingSizeType { get; set; }
        public ShapeSheet.CellData PageInhibitSnap { get; set; }
        public ShapeSheet.CellData PageHeight { get; set; }
        public ShapeSheet.CellData PageScale { get; set; }
        public ShapeSheet.CellData PageWidth { get; set; }
        public ShapeSheet.CellData PageShadowObliqueAngle { get; set; }
        public ShapeSheet.CellData PageShadowOffsetX { get; set; }
        public ShapeSheet.CellData PageShadowOffsetY { get; set; }
        public ShapeSheet.CellData PageShadowScaleFactor { get; set; }
        public ShapeSheet.CellData PageShadowType { get; set; }
        public ShapeSheet.CellData PageUIVisibility { get; set; }
        public ShapeSheet.CellData XGridDensity { get; set; }
        public ShapeSheet.CellData XGridOrigin { get; set; }
        public ShapeSheet.CellData XGridSpacing { get; set; }
        public ShapeSheet.CellData XRulerDensity { get; set; }
        public ShapeSheet.CellData XRulerOrigin { get; set; }
        public ShapeSheet.CellData YGridDensity { get; set; }
        public ShapeSheet.CellData YGridOrigin { get; set; }
        public ShapeSheet.CellData YGridSpacing { get; set; }
        public ShapeSheet.CellData YRulerDensity { get; set; }
        public ShapeSheet.CellData YRulerOrigin { get; set; }
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
        public ShapeSheet.CellData PageDrawingResizeType { get; set; } // new in visio 2010

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.PrintLeftMargin, this.PrintLeftMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintCenterX, this.PrintCenterX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintCenterY, this.PrintCenterY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintOnPage, this.PrintOnPage.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintBottomMargin, this.PrintBottomMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintRightMargin, this.PrintRightMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPagesX, this.PrintPagesX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPagesY, this.PrintPagesY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintTopMargin, this.PrintTopMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPaperKind, this.PrintPaperKind.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintGrid, this.PrintGrid.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPageOrientation, this.PrintPageOrientation.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintScaleX, this.PrintScaleX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintScaleY, this.PrintScaleY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPaperSource, this.PrintPaperSource.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageDrawingScale, this.PageDrawingScale.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageDrawingScaleType, this.PageDrawingScaleType.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageDrawingSizeType, this.PageDrawingSizeType.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageInhibitSnap, this.PageInhibitSnap.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageHeight, this.PageHeight.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageScale, this.PageScale.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageWidth, this.PageWidth.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowObliqueAngle, this.PageShadowObliqueAngle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowOffsetX, this.PageShadowOffsetX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowOffsetY, this.PageShadowOffsetY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowScaleFactor, this.PageShadowScaleFactor.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowType, this.PageShadowType.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageUIVisibility, this.PageUIVisibility.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageDrawingResizeType, this.PageDrawingResizeType.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.XGridDensity, this.XGridDensity.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.XGridOrigin, this.XGridOrigin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.XGridSpacing, this.XGridSpacing.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.XRulerDensity, this.XRulerDensity.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.XRulerOrigin, this.XRulerOrigin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.YGridDensity, this.YGridDensity.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.YGridOrigin, this.YGridOrigin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.YGridSpacing, this.YGridSpacing.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.YRulerDensity, this.YRulerDensity.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.YRulerOrigin, this.YRulerOrigin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutAvenueSizeX, this.PageLayoutAvenueSizeX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutAvenueSizeY, this.PageLayoutAvenueSizeY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutBlockSizeX, this.PageLayoutBlockSizeX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutBlockSizeY, this.PageLayoutBlockSizeY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutCtrlAsInput, this.PageLayoutCtrlAsInput.Formula);
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
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPageShapeSplit, this.PageLayoutPageShapeSplit.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlaceDepth, this.PageLayoutPlaceDepth.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlaceFlip, this.PageLayoutPlaceFlip.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlaceStyle, this.PageLayoutPlaceStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlowCode, this.PageLayoutPlowCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutResizePage, this.PageLayoutResizePage.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutRouteStyle, this.PageLayoutRouteStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutAvoidPageBreaks, this.PageLayoutAvoidPageBreaks.Formula);
            }
        }

        public static PageCells GetCells(IVisio.Shape shape)
        {
            var query = PageCells.lazy_query.Value;
            return query.GetCellGroup(shape);
        }

        private static readonly System.Lazy<PageCellsReader> lazy_query = new System.Lazy<PageCellsReader>();
    }
}