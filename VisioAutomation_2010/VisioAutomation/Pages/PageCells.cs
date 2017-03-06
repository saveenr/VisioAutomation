using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PageCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public ShapeSheet.CellData PageLeftMargin { get; set; }
        public ShapeSheet.CellData CenterX { get; set; }
        public ShapeSheet.CellData CenterY { get; set; }
        public ShapeSheet.CellData OnPage { get; set; }
        public ShapeSheet.CellData PageBottomMargin { get; set; }
        public ShapeSheet.CellData PageRightMargin { get; set; }
        public ShapeSheet.CellData PagesX { get; set; }
        public ShapeSheet.CellData PagesY { get; set; }
        public ShapeSheet.CellData PageTopMargin { get; set; }
        public ShapeSheet.CellData PaperKind { get; set; }
        public ShapeSheet.CellData PrintGrid { get; set; }
        public ShapeSheet.CellData PrintPageOrientation { get; set; }
        public ShapeSheet.CellData ScaleX { get; set; }
        public ShapeSheet.CellData ScaleY { get; set; }
        public ShapeSheet.CellData PaperSource { get; set; }
        public ShapeSheet.CellData DrawingScale { get; set; }
        public ShapeSheet.CellData DrawingScaleType { get; set; }
        public ShapeSheet.CellData DrawingSizeType { get; set; }
        public ShapeSheet.CellData InhibitSnap { get; set; }
        public ShapeSheet.CellData PageHeight { get; set; }
        public ShapeSheet.CellData PageScale { get; set; }
        public ShapeSheet.CellData PageWidth { get; set; }
        public ShapeSheet.CellData ShdwObliqueAngle { get; set; }
        public ShapeSheet.CellData ShdwOffsetX { get; set; }
        public ShapeSheet.CellData ShdwOffsetY { get; set; }
        public ShapeSheet.CellData ShdwScaleFactor { get; set; }
        public ShapeSheet.CellData ShdwType { get; set; }
        public ShapeSheet.CellData UIVisibility { get; set; }
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
        public ShapeSheet.CellData AvenueSizeX { get; set; }
        public ShapeSheet.CellData AvenueSizeY { get; set; }
        public ShapeSheet.CellData BlockSizeX { get; set; }
        public ShapeSheet.CellData BlockSizeY { get; set; }
        public ShapeSheet.CellData CtrlAsInput { get; set; }
        public ShapeSheet.CellData DynamicsOff { get; set; }
        public ShapeSheet.CellData EnableGrid { get; set; }
        public ShapeSheet.CellData LineAdjustFrom { get; set; }
        public ShapeSheet.CellData LineAdjustTo { get; set; }
        public ShapeSheet.CellData LineJumpCode { get; set; }
        public ShapeSheet.CellData LineJumpFactorX { get; set; }
        public ShapeSheet.CellData LineJumpFactorY { get; set; }
        public ShapeSheet.CellData LineJumpStyle { get; set; }
        public ShapeSheet.CellData LineRouteExt { get; set; }
        public ShapeSheet.CellData LineToLineX { get; set; }
        public ShapeSheet.CellData LineToLineY { get; set; }
        public ShapeSheet.CellData LineToNodeX { get; set; }
        public ShapeSheet.CellData LineToNodeY { get; set; }
        public ShapeSheet.CellData PageLineJumpDirX { get; set; }
        public ShapeSheet.CellData PageLineJumpDirY { get; set; }
        public ShapeSheet.CellData PageShapeSplit { get; set; }
        public ShapeSheet.CellData PlaceDepth { get; set; }
        public ShapeSheet.CellData PlaceFlip { get; set; }
        public ShapeSheet.CellData PlaceStyle { get; set; }
        public ShapeSheet.CellData PlowCode { get; set; }
        public ShapeSheet.CellData ResizePage { get; set; }
        public ShapeSheet.CellData RouteStyle { get; set; }
        public ShapeSheet.CellData AvoidPageBreaks { get; set; } // new in visio 2010
        public ShapeSheet.CellData DrawingResizeType { get; set; } // new in visio 2010

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.PrintLeftMargin, this.PageLeftMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintCenterX, this.CenterX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintCenterY, this.CenterY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintOnPage, this.OnPage.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintBottomMargin, this.PageBottomMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintRightMargin, this.PageRightMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPagesX, this.PagesX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPagesY, this.PagesY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintTopMargin, this.PageTopMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPaperKind, this.PaperKind.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintGrid, this.PrintGrid.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPageOrientation, this.PrintPageOrientation.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintScaleX, this.ScaleX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintScaleY, this.ScaleY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPaperSource, this.PaperSource.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageDrawingScale, this.DrawingScale.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageDrawingScaleType, this.DrawingScaleType.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageDrawingSizeType, this.DrawingSizeType.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageInhibitSnap, this.InhibitSnap.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageHeight, this.PageHeight.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageScale, this.PageScale.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageWidth, this.PageWidth.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowObliqueAngle, this.ShdwObliqueAngle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowOffsetX, this.ShdwOffsetX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowOffsetY, this.ShdwOffsetY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowScaleFactor, this.ShdwScaleFactor.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShadowType, this.ShdwType.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageUIVisibility, this.UIVisibility.Formula);
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
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutAvenueSizeX, this.AvenueSizeX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutAvenueSizeY, this.AvenueSizeY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutBlockSizeX, this.BlockSizeX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutBlockSizeY, this.BlockSizeY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutCtrlAsInput, this.CtrlAsInput.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutDynamicsOff, this.DynamicsOff.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutEnableGrid, this.EnableGrid.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineAdjustFrom, this.LineAdjustFrom.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineAdjustTo, this.LineAdjustTo.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpCode, this.LineJumpCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpFactorX, this.LineJumpFactorX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpFactorY, this.LineJumpFactorY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpStyle, this.LineJumpStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineRouteExt, this.LineRouteExt.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineToLineX, this.LineToLineX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineToLineY, this.LineToLineY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineToNodeX, this.LineToNodeX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineToNodeY, this.LineToNodeY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpDirX, this.PageLineJumpDirX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpDirY, this.PageLineJumpDirY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPageShapeSplit, this.PageShapeSplit.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlaceDepth, this.PlaceDepth.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlaceFlip, this.PlaceFlip.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlaceStyle, this.PlaceStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlowCode, this.PlowCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutResizePage, this.ResizePage.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutRouteStyle, this.RouteStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutAvoidPageBreaks, this.AvoidPageBreaks.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageDrawingResizeType, this.DrawingResizeType.Formula);
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