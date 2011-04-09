using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Layout
{

    public partial class PageCells : VA.ShapeSheet.CellDataGroup
    {
        public VA.ShapeSheet.CellData<double> PageLeftMargin { get; set; }
        public VA.ShapeSheet.CellData<double> CenterX { get; set; }
        public VA.ShapeSheet.CellData<double> CenterY { get; set; }
        public VA.ShapeSheet.CellData<int> OnPage { get; set; }
        public VA.ShapeSheet.CellData<double> PageBottomMargin { get; set; }
        public VA.ShapeSheet.CellData<double> PageRightMargin { get; set; }
        public VA.ShapeSheet.CellData<double> PagesX { get; set; }
        public VA.ShapeSheet.CellData<double> PagesY { get; set; }
        public VA.ShapeSheet.CellData<double> PageTopMargin { get; set; }
        public VA.ShapeSheet.CellData<int> PaperKind { get; set; }
        public VA.ShapeSheet.CellData<int> PrintGrid { get; set; }
        public VA.ShapeSheet.CellData<int> PrintPageOrientation { get; set; }
        public VA.ShapeSheet.CellData<double> ScaleX { get; set; }
        public VA.ShapeSheet.CellData<double> ScaleY { get; set; }
        public VA.ShapeSheet.CellData<int> PaperSource { get; set; }
        public VA.ShapeSheet.CellData<double> DrawingScale { get; set; }
        public VA.ShapeSheet.CellData<int> DrawingScaleType { get; set; }
        public VA.ShapeSheet.CellData<int> DrawingSizeType { get; set; }
        public VA.ShapeSheet.CellData<int> InhibitSnap { get; set; }
        public VA.ShapeSheet.CellData<double> PageHeight { get; set; }
        public VA.ShapeSheet.CellData<double> PageScale { get; set; }
        public VA.ShapeSheet.CellData<double> PageWidth { get; set; }
        public VA.ShapeSheet.CellData<double> ShdwObliqueAngle { get; set; }
        public VA.ShapeSheet.CellData<double> ShdwOffsetX { get; set; }
        public VA.ShapeSheet.CellData<double> ShdwOffsetY { get; set; }
        public VA.ShapeSheet.CellData<double> ShdwScaleFactor { get; set; }
        public VA.ShapeSheet.CellData<int> ShdwType { get; set; }
        public VA.ShapeSheet.CellData<double> UIVisibility { get; set; }
        public VA.ShapeSheet.CellData<double> XGridDensity { get; set; }
        public VA.ShapeSheet.CellData<double> XGridOrigin { get; set; }
        public VA.ShapeSheet.CellData<double> XGridSpacing { get; set; }
        public VA.ShapeSheet.CellData<double> XRulerDensity { get; set; }
        public VA.ShapeSheet.CellData<double> XRulerOrigin { get; set; }
        public VA.ShapeSheet.CellData<double> YGridDensity { get; set; }
        public VA.ShapeSheet.CellData<double> YGridOrigin { get; set; }
        public VA.ShapeSheet.CellData<double> YGridSpacing { get; set; }
        public VA.ShapeSheet.CellData<double> YRulerDensity { get; set; }
        public VA.ShapeSheet.CellData<double> YRulerOrigin { get; set; }
        public VA.ShapeSheet.CellData<double> AvenueSizeX { get; set; }
        public VA.ShapeSheet.CellData<double> AvenueSizeY { get; set; }
        public VA.ShapeSheet.CellData<double> BlockSizeX { get; set; }
        public VA.ShapeSheet.CellData<double> BlockSizeY { get; set; }
        public VA.ShapeSheet.CellData<int> CtrlAsInput { get; set; }
        public VA.ShapeSheet.CellData<int> DynamicsOff { get; set; }
        public VA.ShapeSheet.CellData<int> EnableGrid { get; set; }
        public VA.ShapeSheet.CellData<int> LineAdjustFrom { get; set; }
        public VA.ShapeSheet.CellData<double> LineAdjustTo { get; set; }
        public VA.ShapeSheet.CellData<double> LineJumpCode { get; set; }
        public VA.ShapeSheet.CellData<double> LineJumpFactorX { get; set; }
        public VA.ShapeSheet.CellData<double> LineJumpFactorY { get; set; }
        public VA.ShapeSheet.CellData<int> LineJumpStyle { get; set; }
        public VA.ShapeSheet.CellData<double> LineRouteExt { get; set; }
        public VA.ShapeSheet.CellData<double> LineToLineX { get; set; }
        public VA.ShapeSheet.CellData<double> LineToLineY { get; set; }
        public VA.ShapeSheet.CellData<double> LineToNodeX { get; set; }
        public VA.ShapeSheet.CellData<double> LineToNodeY { get; set; }
        public VA.ShapeSheet.CellData<double> PageLineJumpDirX { get; set; }
        public VA.ShapeSheet.CellData<double> PageLineJumpDirY { get; set; }
        public VA.ShapeSheet.CellData<int> PageShapeSplit { get; set; }
        public VA.ShapeSheet.CellData<int> PlaceDepth { get; set; }
        public VA.ShapeSheet.CellData<int> PlaceFlip { get; set; }
        public VA.ShapeSheet.CellData<int> PlaceStyle { get; set; }
        public VA.ShapeSheet.CellData<int> PlowCode { get; set; }
        public VA.ShapeSheet.CellData<int> ResizePage { get; set; }
        public VA.ShapeSheet.CellData<int> RouteStyle { get; set; }

        protected override void _Apply(VA.ShapeSheet.CellDataGroup.ApplyFormula func)
        {
            func(ShapeSheet.SRCConstants.PageLeftMargin, this.PageLeftMargin.Formula);
            func(ShapeSheet.SRCConstants.CenterX, this.CenterX.Formula);
            func(ShapeSheet.SRCConstants.CenterY, this.CenterY.Formula);
            func(ShapeSheet.SRCConstants.OnPage, this.OnPage.Formula);
            func(ShapeSheet.SRCConstants.PageBottomMargin, this.PageBottomMargin.Formula);
            func(ShapeSheet.SRCConstants.PageRightMargin, this.PageRightMargin.Formula);
            func(ShapeSheet.SRCConstants.PagesX, this.PagesX.Formula);
            func(ShapeSheet.SRCConstants.PagesY, this.PagesY.Formula);
            func(ShapeSheet.SRCConstants.PageTopMargin, this.PageTopMargin.Formula);
            func(ShapeSheet.SRCConstants.PaperKind, this.PaperKind.Formula);
            func(ShapeSheet.SRCConstants.PrintGrid, this.PrintGrid.Formula);
            func(ShapeSheet.SRCConstants.PrintPageOrientation, this.PrintPageOrientation.Formula);
            func(ShapeSheet.SRCConstants.ScaleX, this.ScaleX.Formula);
            func(ShapeSheet.SRCConstants.ScaleY, this.ScaleY.Formula);
            func(ShapeSheet.SRCConstants.PaperSource, this.PaperSource.Formula);
            func(ShapeSheet.SRCConstants.DrawingScale, this.DrawingScale.Formula);
            func(ShapeSheet.SRCConstants.DrawingScaleType, this.DrawingScaleType.Formula);
            func(ShapeSheet.SRCConstants.DrawingSizeType, this.DrawingSizeType.Formula);
            func(ShapeSheet.SRCConstants.InhibitSnap, this.InhibitSnap.Formula);
            func(ShapeSheet.SRCConstants.PageHeight, this.PageHeight.Formula);
            func(ShapeSheet.SRCConstants.PageScale, this.PageScale.Formula);
            func(ShapeSheet.SRCConstants.PageWidth, this.PageWidth.Formula);
            func(ShapeSheet.SRCConstants.ShdwObliqueAngle, this.ShdwObliqueAngle.Formula);
            func(ShapeSheet.SRCConstants.ShdwOffsetX, this.ShdwOffsetX.Formula);
            func(ShapeSheet.SRCConstants.ShdwOffsetY, this.ShdwOffsetY.Formula);
            func(ShapeSheet.SRCConstants.ShdwScaleFactor, this.ShdwScaleFactor.Formula);
            func(ShapeSheet.SRCConstants.ShdwType, this.ShdwType.Formula);
            func(ShapeSheet.SRCConstants.UIVisibility, this.UIVisibility.Formula);
            func(ShapeSheet.SRCConstants.XGridDensity, this.XGridDensity.Formula);
            func(ShapeSheet.SRCConstants.XGridOrigin, this.XGridOrigin.Formula);
            func(ShapeSheet.SRCConstants.XGridSpacing, this.XGridSpacing.Formula);
            func(ShapeSheet.SRCConstants.XRulerDensity, this.XRulerDensity.Formula);
            func(ShapeSheet.SRCConstants.XRulerOrigin, this.XRulerOrigin.Formula);
            func(ShapeSheet.SRCConstants.YGridDensity, this.YGridDensity.Formula);
            func(ShapeSheet.SRCConstants.YGridOrigin, this.YGridOrigin.Formula);
            func(ShapeSheet.SRCConstants.YGridSpacing, this.YGridSpacing.Formula);
            func(ShapeSheet.SRCConstants.YRulerDensity, this.YRulerDensity.Formula);
            func(ShapeSheet.SRCConstants.YRulerOrigin, this.YRulerOrigin.Formula);
            func(ShapeSheet.SRCConstants.AvenueSizeX, this.AvenueSizeX.Formula);
            func(ShapeSheet.SRCConstants.AvenueSizeY, this.AvenueSizeY.Formula);
            func(ShapeSheet.SRCConstants.BlockSizeX, this.BlockSizeX.Formula);
            func(ShapeSheet.SRCConstants.BlockSizeY, this.BlockSizeY.Formula);
            func(ShapeSheet.SRCConstants.CtrlAsInput, this.CtrlAsInput.Formula);
            func(ShapeSheet.SRCConstants.DynamicsOff, this.DynamicsOff.Formula);
            func(ShapeSheet.SRCConstants.EnableGrid, this.EnableGrid.Formula);
            func(ShapeSheet.SRCConstants.LineAdjustFrom, this.LineAdjustFrom.Formula);
            func(ShapeSheet.SRCConstants.LineAdjustTo, this.LineAdjustTo.Formula);
            func(ShapeSheet.SRCConstants.LineJumpCode, this.LineJumpCode.Formula);
            func(ShapeSheet.SRCConstants.LineJumpFactorX, this.LineJumpFactorX.Formula);
            func(ShapeSheet.SRCConstants.LineJumpFactorY, this.LineJumpFactorY.Formula);
            func(ShapeSheet.SRCConstants.LineJumpStyle, this.LineJumpStyle.Formula);
            func(ShapeSheet.SRCConstants.LineRouteExt, this.LineRouteExt.Formula);
            func(ShapeSheet.SRCConstants.LineToLineX, this.LineToLineX.Formula);
            func(ShapeSheet.SRCConstants.LineToLineY, this.LineToLineY.Formula);
            func(ShapeSheet.SRCConstants.LineToNodeX, this.LineToNodeX.Formula);
            func(ShapeSheet.SRCConstants.LineToNodeY, this.LineToNodeY.Formula);
            func(ShapeSheet.SRCConstants.PageLineJumpDirX, this.PageLineJumpDirX.Formula);
            func(ShapeSheet.SRCConstants.PageLineJumpDirY, this.PageLineJumpDirY.Formula);
            func(ShapeSheet.SRCConstants.PageShapeSplit, this.PageShapeSplit.Formula);
            func(ShapeSheet.SRCConstants.PlaceDepth, this.PlaceDepth.Formula);
            func(ShapeSheet.SRCConstants.PlaceFlip, this.PlaceFlip.Formula);
            func(ShapeSheet.SRCConstants.PlaceStyle, this.PlaceStyle.Formula);
            func(ShapeSheet.SRCConstants.PlowCode, this.PlowCode.Formula);
            func(ShapeSheet.SRCConstants.ResizePage, this.ResizePage.Formula);
            func(ShapeSheet.SRCConstants.RouteStyle, this.RouteStyle.Formula);
        }

        private static PageCells get_cells_from_row(PageQuery query, VA.ShapeSheet.Query.QueryDataSet<double> qds, int row)
        {
            var cells = new PageCells();
            cells.PageLeftMargin = qds.GetItem(row, query.PageLeftMargin);
            cells.CenterX = qds.GetItem(row, query.CenterX);
            cells.CenterY = qds.GetItem(row, query.CenterY);
            cells.OnPage = qds.GetItem(row, query.OnPage, v => (int)v);
            cells.PageBottomMargin = qds.GetItem(row, query.PageBottomMargin);
            cells.PageRightMargin = qds.GetItem(row, query.PageRightMargin);
            cells.PagesX = qds.GetItem(row, query.PagesX);
            cells.PagesY = qds.GetItem(row, query.PagesY);
            cells.PageTopMargin = qds.GetItem(row, query.PageTopMargin);
            cells.PaperKind = qds.GetItem(row, query.PaperKind, v => (int)v);
            cells.PrintGrid = qds.GetItem(row, query.PrintGrid, v => (int)v);
            cells.PrintPageOrientation = qds.GetItem(row, query.PrintPageOrientation, v => (int)v);
            cells.ScaleX = qds.GetItem(row, query.ScaleX);
            cells.ScaleY = qds.GetItem(row, query.ScaleY);
            cells.PaperSource = qds.GetItem(row, query.PaperSource, v => (int)v);
            cells.DrawingScale = qds.GetItem(row, query.DrawingScale);
            cells.DrawingScaleType = qds.GetItem(row, query.DrawingScaleType, v => (int)v);
            cells.DrawingSizeType = qds.GetItem(row, query.DrawingSizeType, v => (int)v);
            cells.InhibitSnap = qds.GetItem(row, query.InhibitSnap, v => (int)v);
            cells.PageHeight = qds.GetItem(row, query.PageHeight);
            cells.PageScale = qds.GetItem(row, query.PageScale);
            cells.PageWidth = qds.GetItem(row, query.PageWidth);
            cells.ShdwObliqueAngle = qds.GetItem(row, query.ShdwObliqueAngle);
            cells.ShdwOffsetX = qds.GetItem(row, query.ShdwOffsetX);
            cells.ShdwOffsetY = qds.GetItem(row, query.ShdwOffsetY);
            cells.ShdwScaleFactor = qds.GetItem(row, query.ShdwScaleFactor);
            cells.ShdwType = qds.GetItem(row, query.ShdwType, v => (int)v);
            cells.UIVisibility = qds.GetItem(row, query.UIVisibility);
            cells.XGridDensity = qds.GetItem(row, query.XGridDensity);
            cells.XGridOrigin = qds.GetItem(row, query.XGridOrigin);
            cells.XGridSpacing = qds.GetItem(row, query.XGridSpacing);
            cells.XRulerDensity = qds.GetItem(row, query.XRulerDensity);
            cells.XRulerOrigin = qds.GetItem(row, query.XRulerOrigin);
            cells.YGridDensity = qds.GetItem(row, query.YGridDensity);
            cells.YGridOrigin = qds.GetItem(row, query.YGridOrigin);
            cells.YGridSpacing = qds.GetItem(row, query.YGridSpacing);
            cells.YRulerDensity = qds.GetItem(row, query.YRulerDensity);
            cells.YRulerOrigin = qds.GetItem(row, query.YRulerOrigin);
            cells.AvenueSizeX = qds.GetItem(row, query.AvenueSizeX);
            cells.AvenueSizeY = qds.GetItem(row, query.AvenueSizeY);
            cells.BlockSizeX = qds.GetItem(row, query.BlockSizeX);
            cells.BlockSizeY = qds.GetItem(row, query.BlockSizeY);
            cells.CtrlAsInput = qds.GetItem(row, query.CtrlAsInput, v => (int)v);
            cells.DynamicsOff = qds.GetItem(row, query.DynamicsOff, v => (int)v);
            cells.EnableGrid = qds.GetItem(row, query.EnableGrid, v => (int)v);
            cells.LineAdjustFrom = qds.GetItem(row, query.LineAdjustFrom, v => (int)v);
            cells.LineAdjustTo = qds.GetItem(row, query.LineAdjustTo);
            cells.LineJumpCode = qds.GetItem(row, query.LineJumpCode);
            cells.LineJumpFactorX = qds.GetItem(row, query.LineJumpFactorX);
            cells.LineJumpFactorY = qds.GetItem(row, query.LineJumpFactorY);
            cells.LineJumpStyle = qds.GetItem(row, query.LineJumpStyle, v => (int)v);
            cells.LineRouteExt = qds.GetItem(row, query.LineRouteExt);
            cells.LineToLineX = qds.GetItem(row, query.LineToLineX);
            cells.LineToLineY = qds.GetItem(row, query.LineToLineY);
            cells.LineToNodeX = qds.GetItem(row, query.LineToNodeX);
            cells.LineToNodeY = qds.GetItem(row, query.LineToNodeY);
            cells.PageLineJumpDirX = qds.GetItem(row, query.PageLineJumpDirX);
            cells.PageLineJumpDirY = qds.GetItem(row, query.PageLineJumpDirY);
            cells.PageShapeSplit = qds.GetItem(row, query.PageShapeSplit, v => (int)v);
            cells.PlaceDepth = qds.GetItem(row, query.PlaceDepth, v => (int)v);
            cells.PlaceFlip = qds.GetItem(row, query.PlaceFlip, v => (int)v);
            cells.PlaceStyle = qds.GetItem(row, query.PlaceStyle, v => (int)v);
            cells.PlowCode = qds.GetItem(row, query.PlowCode, v => (int)v);
            cells.ResizePage = qds.GetItem(row, query.ResizePage, v => (int)v);
            cells.RouteStyle = qds.GetItem(row, query.RouteStyle, v => (int)v);
            return cells;
        }

        internal static IList<PageCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = new PageQuery();
            return VA.ShapeSheet.CellDataGroup._GetCells(page, shapeids, query, get_cells_from_row);
        }

        internal static PageCells GetCells(IVisio.Shape shape)
        {
            var query = new PageQuery();
            return VA.ShapeSheet.CellDataGroup._GetCells(shape, query, get_cells_from_row);
        }

    }

}