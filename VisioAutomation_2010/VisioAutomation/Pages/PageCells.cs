using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using VisioAutomation.Extensions;
using VAQUERY = VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Pages
{
    public class PageCells : ShapeSheet.CellGroups.CellGroup
    {
        public ShapeSheet.CellData<double> PageLeftMargin { get; set; }
        public ShapeSheet.CellData<double> CenterX { get; set; }
        public ShapeSheet.CellData<double> CenterY { get; set; }
        public ShapeSheet.CellData<int> OnPage { get; set; }
        public ShapeSheet.CellData<double> PageBottomMargin { get; set; }
        public ShapeSheet.CellData<double> PageRightMargin { get; set; }
        public ShapeSheet.CellData<double> PagesX { get; set; }
        public ShapeSheet.CellData<double> PagesY { get; set; }
        public ShapeSheet.CellData<double> PageTopMargin { get; set; }
        public ShapeSheet.CellData<int> PaperKind { get; set; }
        public ShapeSheet.CellData<int> PrintGrid { get; set; }
        public ShapeSheet.CellData<int> PrintPageOrientation { get; set; }
        public ShapeSheet.CellData<double> ScaleX { get; set; }
        public ShapeSheet.CellData<double> ScaleY { get; set; }
        public ShapeSheet.CellData<int> PaperSource { get; set; }
        public ShapeSheet.CellData<double> DrawingScale { get; set; }
        public ShapeSheet.CellData<int> DrawingScaleType { get; set; }
        public ShapeSheet.CellData<int> DrawingSizeType { get; set; }
        public ShapeSheet.CellData<int> InhibitSnap { get; set; }
        public ShapeSheet.CellData<double> PageHeight { get; set; }
        public ShapeSheet.CellData<double> PageScale { get; set; }
        public ShapeSheet.CellData<double> PageWidth { get; set; }
        public ShapeSheet.CellData<double> ShdwObliqueAngle { get; set; }
        public ShapeSheet.CellData<double> ShdwOffsetX { get; set; }
        public ShapeSheet.CellData<double> ShdwOffsetY { get; set; }
        public ShapeSheet.CellData<double> ShdwScaleFactor { get; set; }
        public ShapeSheet.CellData<int> ShdwType { get; set; }
        public ShapeSheet.CellData<double> UIVisibility { get; set; }
        public ShapeSheet.CellData<double> XGridDensity { get; set; }
        public ShapeSheet.CellData<double> XGridOrigin { get; set; }
        public ShapeSheet.CellData<double> XGridSpacing { get; set; }
        public ShapeSheet.CellData<double> XRulerDensity { get; set; }
        public ShapeSheet.CellData<double> XRulerOrigin { get; set; }
        public ShapeSheet.CellData<double> YGridDensity { get; set; }
        public ShapeSheet.CellData<double> YGridOrigin { get; set; }
        public ShapeSheet.CellData<double> YGridSpacing { get; set; }
        public ShapeSheet.CellData<double> YRulerDensity { get; set; }
        public ShapeSheet.CellData<double> YRulerOrigin { get; set; }
        public ShapeSheet.CellData<double> AvenueSizeX { get; set; }
        public ShapeSheet.CellData<double> AvenueSizeY { get; set; }
        public ShapeSheet.CellData<double> BlockSizeX { get; set; }
        public ShapeSheet.CellData<double> BlockSizeY { get; set; }
        public ShapeSheet.CellData<int> CtrlAsInput { get; set; }
        public ShapeSheet.CellData<int> DynamicsOff { get; set; }
        public ShapeSheet.CellData<int> EnableGrid { get; set; }
        public ShapeSheet.CellData<int> LineAdjustFrom { get; set; }
        public ShapeSheet.CellData<double> LineAdjustTo { get; set; }
        public ShapeSheet.CellData<double> LineJumpCode { get; set; }
        public ShapeSheet.CellData<double> LineJumpFactorX { get; set; }
        public ShapeSheet.CellData<double> LineJumpFactorY { get; set; }
        public ShapeSheet.CellData<int> LineJumpStyle { get; set; }
        public ShapeSheet.CellData<double> LineRouteExt { get; set; }
        public ShapeSheet.CellData<double> LineToLineX { get; set; }
        public ShapeSheet.CellData<double> LineToLineY { get; set; }
        public ShapeSheet.CellData<double> LineToNodeX { get; set; }
        public ShapeSheet.CellData<double> LineToNodeY { get; set; }
        public ShapeSheet.CellData<double> PageLineJumpDirX { get; set; }
        public ShapeSheet.CellData<double> PageLineJumpDirY { get; set; }
        public ShapeSheet.CellData<int> PageShapeSplit { get; set; }
        public ShapeSheet.CellData<int> PlaceDepth { get; set; }
        public ShapeSheet.CellData<int> PlaceFlip { get; set; }
        public ShapeSheet.CellData<int> PlaceStyle { get; set; }
        public ShapeSheet.CellData<int> PlowCode { get; set; }
        public ShapeSheet.CellData<int> ResizePage { get; set; }
        public ShapeSheet.CellData<int> RouteStyle { get; set; }
        public ShapeSheet.CellData<int> AvoidPageBreaks { get; set; } // new in visio 2010
        public ShapeSheet.CellData<int> DrawingResizeType { get; set; } // new in visio 2010

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.PageLeftMargin, this.PageLeftMargin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CenterX, this.CenterX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CenterY, this.CenterY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.OnPage, this.OnPage.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PageBottomMargin, this.PageBottomMargin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PageRightMargin, this.PageRightMargin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PagesX, this.PagesX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PagesY, this.PagesY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PageTopMargin, this.PageTopMargin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PaperKind, this.PaperKind.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PrintGrid, this.PrintGrid.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PrintPageOrientation, this.PrintPageOrientation.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ScaleX, this.ScaleX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ScaleY, this.ScaleY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PaperSource, this.PaperSource.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.DrawingScale, this.DrawingScale.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.DrawingScaleType, this.DrawingScaleType.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.DrawingSizeType, this.DrawingSizeType.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.InhibitSnap, this.InhibitSnap.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PageHeight, this.PageHeight.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PageScale, this.PageScale.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PageWidth, this.PageWidth.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwObliqueAngle, this.ShdwObliqueAngle.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwOffsetX, this.ShdwOffsetX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwOffsetY, this.ShdwOffsetY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwScaleFactor, this.ShdwScaleFactor.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ShdwType, this.ShdwType.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.UIVisibility, this.UIVisibility.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.XGridDensity, this.XGridDensity.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.XGridOrigin, this.XGridOrigin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.XGridSpacing, this.XGridSpacing.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.XRulerDensity, this.XRulerDensity.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.XRulerOrigin, this.XRulerOrigin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.YGridDensity, this.YGridDensity.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.YGridOrigin, this.YGridOrigin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.YGridSpacing, this.YGridSpacing.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.YRulerDensity, this.YRulerDensity.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.YRulerOrigin, this.YRulerOrigin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.AvenueSizeX, this.AvenueSizeX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.AvenueSizeY, this.AvenueSizeY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.BlockSizeX, this.BlockSizeX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.BlockSizeY, this.BlockSizeY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.CtrlAsInput, this.CtrlAsInput.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.DynamicsOff, this.DynamicsOff.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.EnableGrid, this.EnableGrid.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineAdjustFrom, this.LineAdjustFrom.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineAdjustTo, this.LineAdjustTo.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineJumpCode, this.LineJumpCode.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineJumpFactorX, this.LineJumpFactorX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineJumpFactorY, this.LineJumpFactorY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineJumpStyle, this.LineJumpStyle.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineRouteExt, this.LineRouteExt.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineToLineX, this.LineToLineX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineToLineY, this.LineToLineY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineToNodeX, this.LineToNodeX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LineToNodeY, this.LineToNodeY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PageLineJumpDirX, this.PageLineJumpDirX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PageLineJumpDirY, this.PageLineJumpDirY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PageShapeSplit, this.PageShapeSplit.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PlaceDepth, this.PlaceDepth.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PlaceFlip, this.PlaceFlip.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PlaceStyle, this.PlaceStyle.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.PlowCode, this.PlowCode.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.ResizePage, this.ResizePage.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.RouteStyle, this.RouteStyle.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.AvoidPageBreaks, this.AvoidPageBreaks.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.DrawingResizeType, this.DrawingResizeType.Formula);
            }
        }

        public static PageCells GetCells(IVisio.Shape shape)
        {
            var query = PageCells.lazy_query.Value;
            return ShapeSheet.CellGroups.CellGroup._GetCells<PageCells, double>(shape, query, query.GetCells);
        }

        private static System.Lazy<ShapeSheet.Query.Common.PageCellQuery> lazy_query = new System.Lazy<ShapeSheet.Query.Common.PageCellQuery>();


    }
}