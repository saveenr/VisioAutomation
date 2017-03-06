using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Pages
{
    class PageCellsReader : SingleRowReader<VisioAutomation.Pages.PageCells>
    {
        public CellColumn PrintLeftMargin { get; set; }
        public CellColumn PrintCenterX { get; set; }
        public CellColumn PrintCenterY { get; set; }
        public CellColumn PrintOnPage { get; set; }
        public CellColumn PrintPageBottomMargin { get; set; }
        public CellColumn PrintRightMargin { get; set; }
        public CellColumn PrintPagesX { get; set; }
        public CellColumn PrintPagesY { get; set; }
        public CellColumn PageTopMargin { get; set; }
        public CellColumn PaperKind { get; set; }
        public CellColumn PrintGrid { get; set; }
        public CellColumn PrintPageOrientation { get; set; }
        public CellColumn ScaleX { get; set; }
        public CellColumn ScaleY { get; set; }
        public CellColumn PaperSource { get; set; }
        public CellColumn DrawingScale { get; set; }
        public CellColumn DrawingScaleType { get; set; }
        public CellColumn DrawingSizeType { get; set; }
        public CellColumn InhibitSnap { get; set; }
        public CellColumn PageHeight { get; set; }
        public CellColumn PageScale { get; set; }
        public CellColumn PageWidth { get; set; }
        public CellColumn ShdwObliqueAngle { get; set; }
        public CellColumn ShdwOffsetX { get; set; }
        public CellColumn ShdwOffsetY { get; set; }
        public CellColumn ShdwScaleFactor { get; set; }
        public CellColumn ShdwType { get; set; }
        public CellColumn UIVisibility { get; set; }
        public CellColumn XGridDensity { get; set; }
        public CellColumn XGridOrigin { get; set; }
        public CellColumn XGridSpacing { get; set; }
        public CellColumn XRulerDensity { get; set; }
        public CellColumn XRulerOrigin { get; set; }
        public CellColumn YGridDensity { get; set; }
        public CellColumn YGridOrigin { get; set; }
        public CellColumn YGridSpacing { get; set; }
        public CellColumn YRulerDensity { get; set; }
        public CellColumn YRulerOrigin { get; set; }
        public CellColumn AvenueSizeX { get; set; }
        public CellColumn AvenueSizeY { get; set; }
        public CellColumn BlockSizeX { get; set; }
        public CellColumn BlockSizeY { get; set; }
        public CellColumn CtrlAsInput { get; set; }
        public CellColumn DynamicsOff { get; set; }
        public CellColumn EnableGrid { get; set; }
        public CellColumn LineAdjustFrom { get; set; }
        public CellColumn LineAdjustTo { get; set; }
        public CellColumn LineJumpCode { get; set; }
        public CellColumn LineJumpFactorX { get; set; }
        public CellColumn LineJumpFactorY { get; set; }
        public CellColumn LineJumpStyle { get; set; }
        public CellColumn LineRouteExt { get; set; }
        public CellColumn LineToLineX { get; set; }
        public CellColumn LineToLineY { get; set; }
        public CellColumn LineToNodeX { get; set; }
        public CellColumn LineToNodeY { get; set; }
        public CellColumn PageLineJumpDirX { get; set; }
        public CellColumn PageLineJumpDirY { get; set; }
        public CellColumn PageShapeSplit { get; set; }
        public CellColumn PlaceDepth { get; set; }
        public CellColumn PlaceFlip { get; set; }
        public CellColumn PlaceStyle { get; set; }
        public CellColumn PlowCode { get; set; }
        public CellColumn ResizePage { get; set; }
        public CellColumn RouteStyle { get; set; }
        public CellColumn AvoidPageBreaks { get; set; }
        public CellColumn DrawingResizeType { get; set; }

        public PageCellsReader()
        {
            this.PrintLeftMargin = this.query.AddCell(SrcConstants.PrintLeftMargin, nameof(SrcConstants.PrintLeftMargin));
            this.PrintCenterX = this.query.AddCell(SrcConstants.PrintCenterX, nameof(SrcConstants.PrintCenterX));
            this.PrintCenterY = this.query.AddCell(SrcConstants.PrintCenterY, nameof(SrcConstants.PrintCenterY));
            this.PrintOnPage = this.query.AddCell(SrcConstants.PrintOnPage, nameof(SrcConstants.PrintOnPage));
            this.PrintPageBottomMargin = this.query.AddCell(SrcConstants.PrintBottomMargin, nameof(SrcConstants.PrintBottomMargin));
            this.PrintRightMargin = this.query.AddCell(SrcConstants.PrintRightMargin, nameof(SrcConstants.PrintRightMargin));
            this.PrintPagesX = this.query.AddCell(SrcConstants.PrintPagesX, nameof(SrcConstants.PrintPagesX));
            this.PrintPagesY = this.query.AddCell(SrcConstants.PrintPagesY, nameof(SrcConstants.PrintPagesY));
            this.PageTopMargin = this.query.AddCell(SrcConstants.PrintTopMargin, nameof(SrcConstants.PrintTopMargin));
            this.PaperKind = this.query.AddCell(SrcConstants.PrintPaperKind, nameof(SrcConstants.PrintPaperKind));
            this.PrintGrid = this.query.AddCell(SrcConstants.PrintGrid, nameof(SrcConstants.PrintGrid));
            this.PrintPageOrientation = this.query.AddCell(SrcConstants.PrintPageOrientation, nameof(SrcConstants.PrintPageOrientation));
            this.ScaleX = this.query.AddCell(SrcConstants.PrintScaleX, nameof(SrcConstants.PrintScaleX));
            this.ScaleY = this.query.AddCell(SrcConstants.PrintScaleY, nameof(SrcConstants.PrintScaleY));
            this.PaperSource = this.query.AddCell(SrcConstants.PrintPaperSource, nameof(SrcConstants.PrintPaperSource));
            this.DrawingScale = this.query.AddCell(SrcConstants.PageDrawingScale, nameof(SrcConstants.PageDrawingScale));
            this.DrawingScaleType = this.query.AddCell(SrcConstants.PageDrawingScaleType, nameof(SrcConstants.PageDrawingScaleType));
            this.DrawingSizeType = this.query.AddCell(SrcConstants.PageDrawingSizeType, nameof(SrcConstants.PageDrawingSizeType));
            this.InhibitSnap = this.query.AddCell(SrcConstants.PageInhibitSnap, nameof(SrcConstants.PageInhibitSnap));
            this.PageHeight = this.query.AddCell(SrcConstants.PageHeight, nameof(SrcConstants.PageHeight));
            this.PageScale = this.query.AddCell(SrcConstants.PageScale, nameof(SrcConstants.PageScale));
            this.PageWidth = this.query.AddCell(SrcConstants.PageWidth, nameof(SrcConstants.PageWidth));
            this.ShdwObliqueAngle = this.query.AddCell(SrcConstants.PageShadowObliqueAngle, nameof(SrcConstants.PageShadowObliqueAngle));
            this.ShdwOffsetX = this.query.AddCell(SrcConstants.PageShadowOffsetX, nameof(SrcConstants.PageShadowOffsetX));
            this.ShdwOffsetY = this.query.AddCell(SrcConstants.PageShadowOffsetY, nameof(SrcConstants.PageShadowOffsetY));
            this.ShdwScaleFactor = this.query.AddCell(SrcConstants.PageShadowScaleFactor, nameof(SrcConstants.PageShadowScaleFactor));
            this.ShdwType = this.query.AddCell(SrcConstants.PageShadowType, nameof(SrcConstants.PageShadowType));
            this.UIVisibility = this.query.AddCell(SrcConstants.PageUIVisibility, nameof(SrcConstants.PageUIVisibility));
            this.XGridDensity = this.query.AddCell(SrcConstants.XGridDensity, nameof(SrcConstants.XGridDensity));
            this.XGridOrigin = this.query.AddCell(SrcConstants.XGridOrigin, nameof(SrcConstants.XGridOrigin));
            this.XGridSpacing = this.query.AddCell(SrcConstants.XGridSpacing, nameof(SrcConstants.XGridSpacing));
            this.XRulerDensity = this.query.AddCell(SrcConstants.XRulerDensity, nameof(SrcConstants.XRulerDensity));
            this.XRulerOrigin = this.query.AddCell(SrcConstants.XRulerOrigin, nameof(SrcConstants.XRulerOrigin));
            this.YGridDensity = this.query.AddCell(SrcConstants.YGridDensity, nameof(SrcConstants.YGridDensity));
            this.YGridOrigin = this.query.AddCell(SrcConstants.YGridOrigin, nameof(SrcConstants.YGridOrigin));
            this.YGridSpacing = this.query.AddCell(SrcConstants.YGridSpacing, nameof(SrcConstants.YGridSpacing));
            this.YRulerDensity = this.query.AddCell(SrcConstants.YRulerDensity, nameof(SrcConstants.YRulerDensity));
            this.YRulerOrigin = this.query.AddCell(SrcConstants.YRulerOrigin, nameof(SrcConstants.YRulerOrigin));
            this.AvenueSizeX = this.query.AddCell(SrcConstants.PageLayoutAvenueSizeX, nameof(SrcConstants.PageLayoutAvenueSizeX));
            this.AvenueSizeY = this.query.AddCell(SrcConstants.PageLayoutAvenueSizeY, nameof(SrcConstants.PageLayoutAvenueSizeY));
            this.BlockSizeX = this.query.AddCell(SrcConstants.PageLayoutBlockSizeX, nameof(SrcConstants.PageLayoutBlockSizeX));
            this.BlockSizeY = this.query.AddCell(SrcConstants.PageLayoutBlockSizeY, nameof(SrcConstants.PageLayoutBlockSizeY));
            this.CtrlAsInput = this.query.AddCell(SrcConstants.PageLayoutCtrlAsInput, nameof(SrcConstants.PageLayoutCtrlAsInput));
            this.DynamicsOff = this.query.AddCell(SrcConstants.PageLayoutDynamicsOff, nameof(SrcConstants.PageLayoutDynamicsOff));
            this.EnableGrid = this.query.AddCell(SrcConstants.PageLayoutEnableGrid, nameof(SrcConstants.PageLayoutEnableGrid));
            this.LineAdjustFrom = this.query.AddCell(SrcConstants.PageLayoutLineAdjustFrom, nameof(SrcConstants.PageLayoutLineAdjustFrom));
            this.LineAdjustTo = this.query.AddCell(SrcConstants.PageLayoutLineAdjustTo, nameof(SrcConstants.PageLayoutLineAdjustTo));
            this.LineJumpCode = this.query.AddCell(SrcConstants.PageLayoutLineJumpCode, nameof(SrcConstants.PageLayoutLineJumpCode));
            this.LineJumpFactorX = this.query.AddCell(SrcConstants.PageLayoutLineJumpFactorX, nameof(SrcConstants.PageLayoutLineJumpFactorX));
            this.LineJumpFactorY = this.query.AddCell(SrcConstants.PageLayoutLineJumpFactorY, nameof(SrcConstants.PageLayoutLineJumpFactorY));
            this.LineJumpStyle = this.query.AddCell(SrcConstants.PageLayoutLineJumpStyle, nameof(SrcConstants.PageLayoutLineJumpStyle));
            this.LineRouteExt = this.query.AddCell(SrcConstants.PageLayoutLineRouteExt, nameof(SrcConstants.PageLayoutLineRouteExt));
            this.LineToLineX = this.query.AddCell(SrcConstants.PageLayoutLineToLineX, nameof(SrcConstants.PageLayoutLineToLineX));
            this.LineToLineY = this.query.AddCell(SrcConstants.PageLayoutLineToLineY, nameof(SrcConstants.PageLayoutLineToLineY));
            this.LineToNodeX = this.query.AddCell(SrcConstants.PageLayoutLineToNodeX, nameof(SrcConstants.PageLayoutLineToNodeX));
            this.LineToNodeY = this.query.AddCell(SrcConstants.PageLayoutLineToNodeY, nameof(SrcConstants.PageLayoutLineToNodeY));
            this.PageLineJumpDirX = this.query.AddCell(SrcConstants.PageLayoutLineJumpDirX, nameof(SrcConstants.PageLayoutLineJumpDirX));
            this.PageLineJumpDirY = this.query.AddCell(SrcConstants.PageLayoutLineJumpDirY, nameof(SrcConstants.PageLayoutLineJumpDirY));
            this.PageShapeSplit = this.query.AddCell(SrcConstants.PageLayoutPageShapeSplit, nameof(SrcConstants.PageLayoutPageShapeSplit));
            this.PlaceDepth = this.query.AddCell(SrcConstants.PageLayoutPlaceDepth, nameof(SrcConstants.PageLayoutPlaceDepth));
            this.PlaceFlip = this.query.AddCell(SrcConstants.PageLayoutPlaceFlip, nameof(SrcConstants.PageLayoutPlaceFlip));
            this.PlaceStyle = this.query.AddCell(SrcConstants.PageLayoutPlaceStyle, nameof(SrcConstants.PageLayoutPlaceStyle));
            this.PlowCode = this.query.AddCell(SrcConstants.PageLayoutPlowCode, nameof(SrcConstants.PageLayoutPlowCode));
            this.ResizePage = this.query.AddCell(SrcConstants.PageLayoutResizePage, nameof(SrcConstants.PageLayoutResizePage));
            this.RouteStyle = this.query.AddCell(SrcConstants.PageLayoutRouteStyle, nameof(SrcConstants.PageLayoutRouteStyle));
            this.AvoidPageBreaks = this.query.AddCell(SrcConstants.PageLayoutAvoidPageBreaks, nameof(SrcConstants.PageLayoutAvoidPageBreaks));

            // Page cells
            this.DrawingResizeType = this.query.AddCell(SrcConstants.PageDrawingResizeType, nameof(SrcConstants.PageDrawingResizeType));
        }


        public override Pages.PageCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Pages.PageCells();
            cells.PrintLeftMargin = row[this.PrintLeftMargin];
            cells.PrintCenterX = row[this.PrintCenterX];
            cells.PrintCenterY = row[this.PrintCenterY];
            cells.PrintOnPage = row[this.PrintOnPage];
            cells.PrintBottomMargin = row[this.PrintPageBottomMargin];
            cells.PrintRightMargin = row[this.PrintRightMargin];
            cells.PrintPagesX = row[this.PrintPagesX];
            cells.PrintPagesY = row[this.PrintPagesY];
            cells.PrintTopMargin = row[this.PageTopMargin];
            cells.PrintPaperKind = row[this.PaperKind];
            cells.PrintGrid = row[this.PrintGrid];
            cells.PrintPageOrientation = row[this.PrintPageOrientation];
            cells.PrintScaleX = row[this.ScaleX];
            cells.PrintScaleY = row[this.ScaleY];
            cells.PrintPaperSource = row[this.PaperSource];
            cells.DrawingScale = row[this.DrawingScale];
            cells.DrawingScaleType = row[this.DrawingScaleType];
            cells.DrawingSizeType = row[this.DrawingSizeType];
            cells.InhibitSnap = row[this.InhibitSnap];
            cells.PageHeight = row[this.PageHeight];
            cells.PageScale = row[this.PageScale];
            cells.PageWidth = row[this.PageWidth];
            cells.ShdwObliqueAngle = row[this.ShdwObliqueAngle];
            cells.ShdwOffsetX = row[this.ShdwOffsetX];
            cells.ShdwOffsetY = row[this.ShdwOffsetY];
            cells.ShdwScaleFactor = row[this.ShdwScaleFactor];
            cells.ShdwType = row[this.ShdwType];
            cells.UIVisibility = row[this.UIVisibility];
            cells.XGridDensity = row[this.XGridDensity];
            cells.XGridOrigin = row[this.XGridOrigin];
            cells.XGridSpacing = row[this.XGridSpacing];
            cells.XRulerDensity = row[this.XRulerDensity];
            cells.XRulerOrigin = row[this.XRulerOrigin];
            cells.YGridDensity = row[this.YGridDensity];
            cells.YGridOrigin = row[this.YGridOrigin];
            cells.YGridSpacing = row[this.YGridSpacing];
            cells.YRulerDensity = row[this.YRulerDensity];
            cells.YRulerOrigin = row[this.YRulerOrigin];
            cells.PageLayoutAvenueSizeX = row[this.AvenueSizeX];
            cells.PageLayoutAvenueSizeY = row[this.AvenueSizeY];
            cells.PageLayoutBlockSizeX = row[this.BlockSizeX];
            cells.PageLayoutBlockSizeY = row[this.BlockSizeY];
            cells.PageLayoutCtrlAsInput = row[this.CtrlAsInput];
            cells.PageLayoutDynamicsOff = row[this.DynamicsOff];
            cells.PageLayoutEnableGrid = row[this.EnableGrid];
            cells.PageLayoutLineAdjustFrom = row[this.LineAdjustFrom];
            cells.PageLayoutLineAdjustTo = row[this.LineAdjustTo];
            cells.PageLayoutLineJumpCode = row[this.LineJumpCode];
            cells.PageLayoutLineJumpFactorX = row[this.LineJumpFactorX];
            cells.PageLayoutLineJumpFactorY = row[this.LineJumpFactorY];
            cells.PageLayoutLineJumpStyle = row[this.LineJumpStyle];
            cells.PageLayoutLineRouteExt = row[this.LineRouteExt];
            cells.PageLayoutLineToLineX = row[this.LineToLineX];
            cells.PageLayoutLineToLineY = row[this.LineToLineY];
            cells.PageLayoutLineToNodeX = row[this.LineToNodeX];
            cells.PageLayoutLineToNodeY = row[this.LineToNodeY];
            cells.PageLayoutLineJumpDirX = row[this.PageLineJumpDirX];
            cells.PageLayoutLineJumpDirY = row[this.PageLineJumpDirY];
            cells.PageLayoutPageShapeSplit = row[this.PageShapeSplit];
            cells.PageLayoutPlaceDepth = row[this.PlaceDepth];
            cells.PageLayoutPlaceFlip = row[this.PlaceFlip];
            cells.PageLayoutPlaceStyle = row[this.PlaceStyle];
            cells.PageLayoutPlowCode = row[this.PlowCode];
            cells.PageLayoutResizePage = row[this.ResizePage];
            cells.PageLayoutRouteStyle = row[this.RouteStyle];
            cells.PageLayoutAvoidPageBreaks = row[this.AvoidPageBreaks];
            cells.DrawingResizeType = row[this.DrawingResizeType];
            return cells;
        }
    }
}