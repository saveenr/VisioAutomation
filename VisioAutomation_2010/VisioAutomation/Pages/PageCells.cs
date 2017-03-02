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
                yield return this.newpair(ShapeSheet.SrcConstants.PageLeftMargin, this.PageLeftMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.CenterX, this.CenterX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.CenterY, this.CenterY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.OnPage, this.OnPage.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageBottomMargin, this.PageBottomMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageRightMargin, this.PageRightMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PagesX, this.PagesX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PagesY, this.PagesY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageTopMargin, this.PageTopMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PaperKind, this.PaperKind.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintGrid, this.PrintGrid.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PrintPageOrientation, this.PrintPageOrientation.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ScaleX, this.ScaleX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ScaleY, this.ScaleY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PaperSource, this.PaperSource.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.DrawingScale, this.DrawingScale.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.DrawingScaleType, this.DrawingScaleType.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.DrawingSizeType, this.DrawingSizeType.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.InhibitSnap, this.InhibitSnap.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageHeight, this.PageHeight.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageScale, this.PageScale.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageWidth, this.PageWidth.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShdwObliqueAngle, this.ShdwObliqueAngle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShdwOffsetX, this.ShdwOffsetX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShdwOffsetY, this.ShdwOffsetY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShdwScaleFactor, this.ShdwScaleFactor.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ShdwType, this.ShdwType.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.UIVisibility, this.UIVisibility.Formula);
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
                yield return this.newpair(ShapeSheet.SrcConstants.AvenueSizeX, this.AvenueSizeX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.AvenueSizeY, this.AvenueSizeY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.BlockSizeX, this.BlockSizeX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.BlockSizeY, this.BlockSizeY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.CtrlAsInput, this.CtrlAsInput.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.DynamicsOff, this.DynamicsOff.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.EnableGrid, this.EnableGrid.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineAdjustFrom, this.LineAdjustFrom.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineAdjustTo, this.LineAdjustTo.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineJumpCode, this.LineJumpCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineJumpFactorX, this.LineJumpFactorX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineJumpFactorY, this.LineJumpFactorY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineJumpStyle, this.LineJumpStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineRouteExt, this.LineRouteExt.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineToLineX, this.LineToLineX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineToLineY, this.LineToLineY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineToNodeX, this.LineToNodeX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LineToNodeY, this.LineToNodeY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLineJumpDirX, this.PageLineJumpDirX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLineJumpDirY, this.PageLineJumpDirY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageShapeSplit, this.PageShapeSplit.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PlaceDepth, this.PlaceDepth.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PlaceFlip, this.PlaceFlip.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PlaceStyle, this.PlaceStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PlowCode, this.PlowCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ResizePage, this.ResizePage.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.RouteStyle, this.RouteStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.AvoidPageBreaks, this.AvoidPageBreaks.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.DrawingResizeType, this.DrawingResizeType.Formula);
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