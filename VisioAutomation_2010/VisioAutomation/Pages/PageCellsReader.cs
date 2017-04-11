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
        public CellColumn PrintBottomMargin { get; set; }
        public CellColumn PrintRightMargin { get; set; }
        public CellColumn PrintPagesX { get; set; }
        public CellColumn PrintPagesY { get; set; }
        public CellColumn PrintTopMargin { get; set; }
        public CellColumn PrintPaperKind { get; set; }
        public CellColumn PrintGrid { get; set; }
        public CellColumn PrintPageOrientation { get; set; }
        public CellColumn PrintScaleX { get; set; }
        public CellColumn PrintScaleY { get; set; }
        public CellColumn PrintPaperSource { get; set; }
        public CellColumn PageDrawingScale { get; set; }
        public CellColumn PageDrawingScaleType { get; set; }
        public CellColumn PageDrawingSizeType { get; set; }
        public CellColumn PageInhibitSnap { get; set; }
        public CellColumn PageHeight { get; set; }
        public CellColumn PageScale { get; set; }
        public CellColumn PageWidth { get; set; }
        public CellColumn PageShadowObliqueAngle { get; set; }
        public CellColumn PageShadowOffsetX { get; set; }
        public CellColumn PageShadowOffsetY { get; set; }
        public CellColumn PageShadowScaleFactor { get; set; }
        public CellColumn PageShadowType { get; set; }
        public CellColumn PageUIVisibility { get; set; }
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
        public CellColumn PageLayoutAvenueSizeX { get; set; }
        public CellColumn PageLayoutAvenueSizeY { get; set; }
        public CellColumn PageLayoutBlockSizeX { get; set; }
        public CellColumn PageLayoutBlockSizeY { get; set; }
        public CellColumn PageLayoutControlAsInput { get; set; }
        public CellColumn PageLayoutDynamicsOff { get; set; }
        public CellColumn PageLayoutEnableGrid { get; set; }
        public CellColumn PageLayoutLineAdjustFrom { get; set; }
        public CellColumn PageLayoutLineAdjustTo { get; set; }
        public CellColumn PageLayoutLineJumpCode { get; set; }
        public CellColumn PageLayoutLineJumpFactorX { get; set; }
        public CellColumn PageLayoutLineJumpFactorY { get; set; }
        public CellColumn PageLayoutLineJumpStyle { get; set; }
        public CellColumn PageLayoutLineRouteExt { get; set; }
        public CellColumn PageLayoutLineToLineX { get; set; }
        public CellColumn PageLayoutLineToLineY { get; set; }
        public CellColumn PageLayoutLineToNodeX { get; set; }
        public CellColumn PageLayoutLineToNodeY { get; set; }
        public CellColumn PageLayoutLineJumpDirX { get; set; }
        public CellColumn PageLayoutLineJumpDirY { get; set; }
        public CellColumn PageLayoutShapeSplit { get; set; }
        public CellColumn PageLayoutPlaceDepth { get; set; }
        public CellColumn PageLayoutPlaceFlip { get; set; }
        public CellColumn PageLayoutPlaceStyle { get; set; }
        public CellColumn PageLayoutPlowCode { get; set; }
        public CellColumn PageLayoutResizePage { get; set; }
        public CellColumn PageLayoutRouteStyle { get; set; }
        public CellColumn PageLayoutAvoidPageBreaks { get; set; }
        public CellColumn PageDrawingResizeType { get; set; }

        public PageCellsReader()
        {
            this.PrintLeftMargin = this.query.AddCell(SrcConstants.PrintLeftMargin, nameof(SrcConstants.PrintLeftMargin));
            this.PrintCenterX = this.query.AddCell(SrcConstants.PrintCenterX, nameof(SrcConstants.PrintCenterX));
            this.PrintCenterY = this.query.AddCell(SrcConstants.PrintCenterY, nameof(SrcConstants.PrintCenterY));
            this.PrintOnPage = this.query.AddCell(SrcConstants.PrintOnPage, nameof(SrcConstants.PrintOnPage));
            this.PrintBottomMargin = this.query.AddCell(SrcConstants.PrintBottomMargin, nameof(SrcConstants.PrintBottomMargin));
            this.PrintRightMargin = this.query.AddCell(SrcConstants.PrintRightMargin, nameof(SrcConstants.PrintRightMargin));
            this.PrintPagesX = this.query.AddCell(SrcConstants.PrintPagesX, nameof(SrcConstants.PrintPagesX));
            this.PrintPagesY = this.query.AddCell(SrcConstants.PrintPagesY, nameof(SrcConstants.PrintPagesY));
            this.PrintTopMargin = this.query.AddCell(SrcConstants.PrintTopMargin, nameof(SrcConstants.PrintTopMargin));
            this.PrintPaperKind = this.query.AddCell(SrcConstants.PrintPaperKind, nameof(SrcConstants.PrintPaperKind));
            this.PrintGrid = this.query.AddCell(SrcConstants.PrintGrid, nameof(SrcConstants.PrintGrid));
            this.PrintPageOrientation = this.query.AddCell(SrcConstants.PrintPageOrientation, nameof(SrcConstants.PrintPageOrientation));
            this.PrintScaleX = this.query.AddCell(SrcConstants.PrintScaleX, nameof(SrcConstants.PrintScaleX));
            this.PrintScaleY = this.query.AddCell(SrcConstants.PrintScaleY, nameof(SrcConstants.PrintScaleY));
            this.PrintPaperSource = this.query.AddCell(SrcConstants.PrintPaperSource, nameof(SrcConstants.PrintPaperSource));

            this.PageDrawingScale = this.query.AddCell(SrcConstants.PageDrawingScale, nameof(SrcConstants.PageDrawingScale));
            this.PageDrawingScaleType = this.query.AddCell(SrcConstants.PageDrawingScaleType, nameof(SrcConstants.PageDrawingScaleType));
            this.PageDrawingSizeType = this.query.AddCell(SrcConstants.PageDrawingSizeType, nameof(SrcConstants.PageDrawingSizeType));
            this.PageInhibitSnap = this.query.AddCell(SrcConstants.PageInhibitSnap, nameof(SrcConstants.PageInhibitSnap));
            this.PageHeight = this.query.AddCell(SrcConstants.PageHeight, nameof(SrcConstants.PageHeight));
            this.PageScale = this.query.AddCell(SrcConstants.PageScale, nameof(SrcConstants.PageScale));
            this.PageWidth = this.query.AddCell(SrcConstants.PageWidth, nameof(SrcConstants.PageWidth));
            this.PageShadowObliqueAngle = this.query.AddCell(SrcConstants.PageShadowObliqueAngle, nameof(SrcConstants.PageShadowObliqueAngle));
            this.PageShadowOffsetX = this.query.AddCell(SrcConstants.PageShadowOffsetX, nameof(SrcConstants.PageShadowOffsetX));
            this.PageShadowOffsetY = this.query.AddCell(SrcConstants.PageShadowOffsetY, nameof(SrcConstants.PageShadowOffsetY));
            this.PageShadowScaleFactor = this.query.AddCell(SrcConstants.PageShadowScaleFactor, nameof(SrcConstants.PageShadowScaleFactor));
            this.PageShadowType = this.query.AddCell(SrcConstants.PageShadowType, nameof(SrcConstants.PageShadowType));
            this.PageUIVisibility = this.query.AddCell(SrcConstants.PageUIVisibility, nameof(SrcConstants.PageUIVisibility));
            this.PageDrawingResizeType = this.query.AddCell(SrcConstants.PageDrawingResizeType, nameof(SrcConstants.PageDrawingResizeType));

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

            this.PageLayoutAvenueSizeX = this.query.AddCell(SrcConstants.PageLayoutAvenueSizeX, nameof(SrcConstants.PageLayoutAvenueSizeX));
            this.PageLayoutAvenueSizeY = this.query.AddCell(SrcConstants.PageLayoutAvenueSizeY, nameof(SrcConstants.PageLayoutAvenueSizeY));
            this.PageLayoutBlockSizeX = this.query.AddCell(SrcConstants.PageLayoutBlockSizeX, nameof(SrcConstants.PageLayoutBlockSizeX));
            this.PageLayoutBlockSizeY = this.query.AddCell(SrcConstants.PageLayoutBlockSizeY, nameof(SrcConstants.PageLayoutBlockSizeY));
            this.PageLayoutControlAsInput = this.query.AddCell(SrcConstants.PageLayoutControlAsInput, nameof(SrcConstants.PageLayoutControlAsInput));
            this.PageLayoutDynamicsOff = this.query.AddCell(SrcConstants.PageLayoutDynamicsOff, nameof(SrcConstants.PageLayoutDynamicsOff));
            this.PageLayoutEnableGrid = this.query.AddCell(SrcConstants.PageLayoutEnableGrid, nameof(SrcConstants.PageLayoutEnableGrid));
            this.PageLayoutLineAdjustFrom = this.query.AddCell(SrcConstants.PageLayoutLineAdjustFrom, nameof(SrcConstants.PageLayoutLineAdjustFrom));
            this.PageLayoutLineAdjustTo = this.query.AddCell(SrcConstants.PageLayoutLineAdjustTo, nameof(SrcConstants.PageLayoutLineAdjustTo));
            this.PageLayoutLineJumpCode = this.query.AddCell(SrcConstants.PageLayoutLineJumpCode, nameof(SrcConstants.PageLayoutLineJumpCode));
            this.PageLayoutLineJumpFactorX = this.query.AddCell(SrcConstants.PageLayoutLineJumpFactorX, nameof(SrcConstants.PageLayoutLineJumpFactorX));
            this.PageLayoutLineJumpFactorY = this.query.AddCell(SrcConstants.PageLayoutLineJumpFactorY, nameof(SrcConstants.PageLayoutLineJumpFactorY));
            this.PageLayoutLineJumpStyle = this.query.AddCell(SrcConstants.PageLayoutLineJumpStyle, nameof(SrcConstants.PageLayoutLineJumpStyle));
            this.PageLayoutLineRouteExt = this.query.AddCell(SrcConstants.PageLayoutLineRouteExt, nameof(SrcConstants.PageLayoutLineRouteExt));
            this.PageLayoutLineToLineX = this.query.AddCell(SrcConstants.PageLayoutLineToLineX, nameof(SrcConstants.PageLayoutLineToLineX));
            this.PageLayoutLineToLineY = this.query.AddCell(SrcConstants.PageLayoutLineToLineY, nameof(SrcConstants.PageLayoutLineToLineY));
            this.PageLayoutLineToNodeX = this.query.AddCell(SrcConstants.PageLayoutLineToNodeX, nameof(SrcConstants.PageLayoutLineToNodeX));
            this.PageLayoutLineToNodeY = this.query.AddCell(SrcConstants.PageLayoutLineToNodeY, nameof(SrcConstants.PageLayoutLineToNodeY));
            this.PageLayoutLineJumpDirX = this.query.AddCell(SrcConstants.PageLayoutLineJumpDirX, nameof(SrcConstants.PageLayoutLineJumpDirX));
            this.PageLayoutLineJumpDirY = this.query.AddCell(SrcConstants.PageLayoutLineJumpDirY, nameof(SrcConstants.PageLayoutLineJumpDirY));
            this.PageLayoutShapeSplit = this.query.AddCell(SrcConstants.PageLayoutShapeSplit, nameof(SrcConstants.PageLayoutShapeSplit));
            this.PageLayoutPlaceDepth = this.query.AddCell(SrcConstants.PageLayoutPlaceDepth, nameof(SrcConstants.PageLayoutPlaceDepth));
            this.PageLayoutPlaceFlip = this.query.AddCell(SrcConstants.PageLayoutPlaceFlip, nameof(SrcConstants.PageLayoutPlaceFlip));
            this.PageLayoutPlaceStyle = this.query.AddCell(SrcConstants.PageLayoutPlaceStyle, nameof(SrcConstants.PageLayoutPlaceStyle));
            this.PageLayoutPlowCode = this.query.AddCell(SrcConstants.PageLayoutPlowCode, nameof(SrcConstants.PageLayoutPlowCode));
            this.PageLayoutResizePage = this.query.AddCell(SrcConstants.PageLayoutResizePage, nameof(SrcConstants.PageLayoutResizePage));
            this.PageLayoutRouteStyle = this.query.AddCell(SrcConstants.PageLayoutRouteStyle, nameof(SrcConstants.PageLayoutRouteStyle));
            this.PageLayoutAvoidPageBreaks = this.query.AddCell(SrcConstants.PageLayoutAvoidPageBreaks, nameof(SrcConstants.PageLayoutAvoidPageBreaks));
        }


        public override Pages.PageCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Pages.PageCells();
            cells.PrintLeftMargin = row[this.PrintLeftMargin];
            cells.PrintCenterX = row[this.PrintCenterX];
            cells.PrintCenterY = row[this.PrintCenterY];
            cells.PrintOnPage = row[this.PrintOnPage];
            cells.PrintBottomMargin = row[this.PrintBottomMargin];
            cells.PrintRightMargin = row[this.PrintRightMargin];
            cells.PrintPagesX = row[this.PrintPagesX];
            cells.PrintPagesY = row[this.PrintPagesY];
            cells.PrintTopMargin = row[this.PrintTopMargin];
            cells.PrintPaperKind = row[this.PrintPaperKind];
            cells.PrintGrid = row[this.PrintGrid];
            cells.PrintPageOrientation = row[this.PrintPageOrientation];
            cells.PrintScaleX = row[this.PrintScaleX];
            cells.PrintScaleY = row[this.PrintScaleY];
            cells.PrintPaperSource = row[this.PrintPaperSource];
            cells.PageDrawingScale = row[this.PageDrawingScale];
            cells.PageDrawingScaleType = row[this.PageDrawingScaleType];
            cells.PageDrawingSizeType = row[this.PageDrawingSizeType];
            cells.PageInhibitSnap = row[this.PageInhibitSnap];
            cells.PageHeight = row[this.PageHeight];
            cells.PageScale = row[this.PageScale];
            cells.PageWidth = row[this.PageWidth];
            cells.PageShadowObliqueAngle = row[this.PageShadowObliqueAngle];
            cells.PageShadowOffsetX = row[this.PageShadowOffsetX];
            cells.PageShadowOffsetY = row[this.PageShadowOffsetY];
            cells.PageShadowScaleFactor = row[this.PageShadowScaleFactor];
            cells.PageShadowType = row[this.PageShadowType];
            cells.PageUIVisibility = row[this.PageUIVisibility];
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
            cells.PageDrawingResizeType = row[this.PageDrawingResizeType];
            return cells;
        }
    }



    class PageLayoutCellsReader : SingleRowReader<VisioAutomation.Pages.PageLayoutCells>
    {
        public CellColumn PageLayoutAvenueSizeX { get; set; }
        public CellColumn PageLayoutAvenueSizeY { get; set; }
        public CellColumn PageLayoutBlockSizeX { get; set; }
        public CellColumn PageLayoutBlockSizeY { get; set; }
        public CellColumn PageLayoutControlAsInput { get; set; }
        public CellColumn PageLayoutDynamicsOff { get; set; }
        public CellColumn PageLayoutEnableGrid { get; set; }
        public CellColumn PageLayoutLineAdjustFrom { get; set; }
        public CellColumn PageLayoutLineAdjustTo { get; set; }
        public CellColumn PageLayoutLineJumpCode { get; set; }
        public CellColumn PageLayoutLineJumpFactorX { get; set; }
        public CellColumn PageLayoutLineJumpFactorY { get; set; }
        public CellColumn PageLayoutLineJumpStyle { get; set; }
        public CellColumn PageLayoutLineRouteExt { get; set; }
        public CellColumn PageLayoutLineToLineX { get; set; }
        public CellColumn PageLayoutLineToLineY { get; set; }
        public CellColumn PageLayoutLineToNodeX { get; set; }
        public CellColumn PageLayoutLineToNodeY { get; set; }
        public CellColumn PageLayoutLineJumpDirX { get; set; }
        public CellColumn PageLayoutLineJumpDirY { get; set; }
        public CellColumn PageLayoutShapeSplit { get; set; }
        public CellColumn PageLayoutPlaceDepth { get; set; }
        public CellColumn PageLayoutPlaceFlip { get; set; }
        public CellColumn PageLayoutPlaceStyle { get; set; }
        public CellColumn PageLayoutPlowCode { get; set; }
        public CellColumn PageLayoutResizePage { get; set; }
        public CellColumn PageLayoutRouteStyle { get; set; }
        public CellColumn PageLayoutAvoidPageBreaks { get; set; }

        public PageLayoutCellsReader()
        {
            this.PageLayoutAvenueSizeX = this.query.AddCell(SrcConstants.PageLayoutAvenueSizeX, nameof(SrcConstants.PageLayoutAvenueSizeX));
            this.PageLayoutAvenueSizeY = this.query.AddCell(SrcConstants.PageLayoutAvenueSizeY, nameof(SrcConstants.PageLayoutAvenueSizeY));
            this.PageLayoutBlockSizeX = this.query.AddCell(SrcConstants.PageLayoutBlockSizeX, nameof(SrcConstants.PageLayoutBlockSizeX));
            this.PageLayoutBlockSizeY = this.query.AddCell(SrcConstants.PageLayoutBlockSizeY, nameof(SrcConstants.PageLayoutBlockSizeY));
            this.PageLayoutControlAsInput = this.query.AddCell(SrcConstants.PageLayoutControlAsInput, nameof(SrcConstants.PageLayoutControlAsInput));
            this.PageLayoutDynamicsOff = this.query.AddCell(SrcConstants.PageLayoutDynamicsOff, nameof(SrcConstants.PageLayoutDynamicsOff));
            this.PageLayoutEnableGrid = this.query.AddCell(SrcConstants.PageLayoutEnableGrid, nameof(SrcConstants.PageLayoutEnableGrid));
            this.PageLayoutLineAdjustFrom = this.query.AddCell(SrcConstants.PageLayoutLineAdjustFrom, nameof(SrcConstants.PageLayoutLineAdjustFrom));
            this.PageLayoutLineAdjustTo = this.query.AddCell(SrcConstants.PageLayoutLineAdjustTo, nameof(SrcConstants.PageLayoutLineAdjustTo));
            this.PageLayoutLineJumpCode = this.query.AddCell(SrcConstants.PageLayoutLineJumpCode, nameof(SrcConstants.PageLayoutLineJumpCode));
            this.PageLayoutLineJumpFactorX = this.query.AddCell(SrcConstants.PageLayoutLineJumpFactorX, nameof(SrcConstants.PageLayoutLineJumpFactorX));
            this.PageLayoutLineJumpFactorY = this.query.AddCell(SrcConstants.PageLayoutLineJumpFactorY, nameof(SrcConstants.PageLayoutLineJumpFactorY));
            this.PageLayoutLineJumpStyle = this.query.AddCell(SrcConstants.PageLayoutLineJumpStyle, nameof(SrcConstants.PageLayoutLineJumpStyle));
            this.PageLayoutLineRouteExt = this.query.AddCell(SrcConstants.PageLayoutLineRouteExt, nameof(SrcConstants.PageLayoutLineRouteExt));
            this.PageLayoutLineToLineX = this.query.AddCell(SrcConstants.PageLayoutLineToLineX, nameof(SrcConstants.PageLayoutLineToLineX));
            this.PageLayoutLineToLineY = this.query.AddCell(SrcConstants.PageLayoutLineToLineY, nameof(SrcConstants.PageLayoutLineToLineY));
            this.PageLayoutLineToNodeX = this.query.AddCell(SrcConstants.PageLayoutLineToNodeX, nameof(SrcConstants.PageLayoutLineToNodeX));
            this.PageLayoutLineToNodeY = this.query.AddCell(SrcConstants.PageLayoutLineToNodeY, nameof(SrcConstants.PageLayoutLineToNodeY));
            this.PageLayoutLineJumpDirX = this.query.AddCell(SrcConstants.PageLayoutLineJumpDirX, nameof(SrcConstants.PageLayoutLineJumpDirX));
            this.PageLayoutLineJumpDirY = this.query.AddCell(SrcConstants.PageLayoutLineJumpDirY, nameof(SrcConstants.PageLayoutLineJumpDirY));
            this.PageLayoutShapeSplit = this.query.AddCell(SrcConstants.PageLayoutShapeSplit, nameof(SrcConstants.PageLayoutShapeSplit));
            this.PageLayoutPlaceDepth = this.query.AddCell(SrcConstants.PageLayoutPlaceDepth, nameof(SrcConstants.PageLayoutPlaceDepth));
            this.PageLayoutPlaceFlip = this.query.AddCell(SrcConstants.PageLayoutPlaceFlip, nameof(SrcConstants.PageLayoutPlaceFlip));
            this.PageLayoutPlaceStyle = this.query.AddCell(SrcConstants.PageLayoutPlaceStyle, nameof(SrcConstants.PageLayoutPlaceStyle));
            this.PageLayoutPlowCode = this.query.AddCell(SrcConstants.PageLayoutPlowCode, nameof(SrcConstants.PageLayoutPlowCode));
            this.PageLayoutResizePage = this.query.AddCell(SrcConstants.PageLayoutResizePage, nameof(SrcConstants.PageLayoutResizePage));
            this.PageLayoutRouteStyle = this.query.AddCell(SrcConstants.PageLayoutRouteStyle, nameof(SrcConstants.PageLayoutRouteStyle));
            this.PageLayoutAvoidPageBreaks = this.query.AddCell(SrcConstants.PageLayoutAvoidPageBreaks, nameof(SrcConstants.PageLayoutAvoidPageBreaks));
        }


        public override Pages.PageLayoutCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Pages.PageLayoutCells();
            cells.PageLayoutAvenueSizeX = row[this.PageLayoutAvenueSizeX];
            cells.PageLayoutAvenueSizeY = row[this.PageLayoutAvenueSizeY];
            cells.PageLayoutBlockSizeX = row[this.PageLayoutBlockSizeX];
            cells.PageLayoutBlockSizeY = row[this.PageLayoutBlockSizeY];
            cells.PageLayoutCtrlAsInput = row[this.PageLayoutControlAsInput];
            cells.PageLayoutDynamicsOff = row[this.PageLayoutDynamicsOff];
            cells.PageLayoutEnableGrid = row[this.PageLayoutEnableGrid];
            cells.PageLayoutLineAdjustFrom = row[this.PageLayoutLineAdjustFrom];
            cells.PageLayoutLineAdjustTo = row[this.PageLayoutLineAdjustTo];
            cells.PageLayoutLineJumpCode = row[this.PageLayoutLineJumpCode];
            cells.PageLayoutLineJumpFactorX = row[this.PageLayoutLineJumpFactorX];
            cells.PageLayoutLineJumpFactorY = row[this.PageLayoutLineJumpFactorY];
            cells.PageLayoutLineJumpStyle = row[this.PageLayoutLineJumpStyle];
            cells.PageLayoutLineRouteExt = row[this.PageLayoutLineRouteExt];
            cells.PageLayoutLineToLineX = row[this.PageLayoutLineToLineX];
            cells.PageLayoutLineToLineY = row[this.PageLayoutLineToLineY];
            cells.PageLayoutLineToNodeX = row[this.PageLayoutLineToNodeX];
            cells.PageLayoutLineToNodeY = row[this.PageLayoutLineToNodeY];
            cells.PageLayoutLineJumpDirX = row[this.PageLayoutLineJumpDirX];
            cells.PageLayoutLineJumpDirY = row[this.PageLayoutLineJumpDirY];
            cells.PageLayoutPageShapeSplit = row[this.PageLayoutShapeSplit];
            cells.PageLayoutPlaceDepth = row[this.PageLayoutPlaceDepth];
            cells.PageLayoutPlaceFlip = row[this.PageLayoutPlaceFlip];
            cells.PageLayoutPlaceStyle = row[this.PageLayoutPlaceStyle];
            cells.PageLayoutPlowCode = row[this.PageLayoutPlowCode];
            cells.PageLayoutResizePage = row[this.PageLayoutResizePage];
            cells.PageLayoutRouteStyle = row[this.PageLayoutRouteStyle];
            cells.PageLayoutAvoidPageBreaks = row[this.PageLayoutAvoidPageBreaks];
            return cells;
        }
    }

}