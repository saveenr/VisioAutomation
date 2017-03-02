using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Pages
{
    class PageCellsReader : SingleRowReader<VisioAutomation.Pages.PageCells>
    {
        public CellColumn PageLeftMargin { get; set; }
        public CellColumn CenterX { get; set; }
        public CellColumn CenterY { get; set; }
        public CellColumn OnPage { get; set; }
        public CellColumn PageBottomMargin { get; set; }
        public CellColumn PageRightMargin { get; set; }
        public CellColumn PagesX { get; set; }
        public CellColumn PagesY { get; set; }
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
            this.PageLeftMargin = this.query.AddCell(SrcConstants.PageLeftMargin, nameof(SrcConstants.PageLeftMargin));
            this.CenterX = this.query.AddCell(SrcConstants.CenterX, nameof(SrcConstants.CenterX));
            this.CenterY = this.query.AddCell(SrcConstants.CenterY, nameof(SrcConstants.CenterY));
            this.OnPage = this.query.AddCell(SrcConstants.OnPage, nameof(SrcConstants.OnPage));
            this.PageBottomMargin = this.query.AddCell(SrcConstants.PageBottomMargin, nameof(SrcConstants.PageBottomMargin));
            this.PageRightMargin = this.query.AddCell(SrcConstants.PageRightMargin, nameof(SrcConstants.PageRightMargin));
            this.PagesX = this.query.AddCell(SrcConstants.PagesX, nameof(SrcConstants.PagesX));
            this.PagesY = this.query.AddCell(SrcConstants.PagesY, nameof(SrcConstants.PagesY));
            this.PageTopMargin = this.query.AddCell(SrcConstants.PageTopMargin, nameof(SrcConstants.PageTopMargin));
            this.PaperKind = this.query.AddCell(SrcConstants.PaperKind, nameof(SrcConstants.PaperKind));
            this.PrintGrid = this.query.AddCell(SrcConstants.PrintGrid, nameof(SrcConstants.PrintGrid));
            this.PrintPageOrientation = this.query.AddCell(SrcConstants.PrintPageOrientation, nameof(SrcConstants.PrintPageOrientation));
            this.ScaleX = this.query.AddCell(SrcConstants.ScaleX, nameof(SrcConstants.ScaleX));
            this.ScaleY = this.query.AddCell(SrcConstants.ScaleY, nameof(SrcConstants.ScaleY));
            this.PaperSource = this.query.AddCell(SrcConstants.PaperSource, nameof(SrcConstants.PaperSource));
            this.DrawingScale = this.query.AddCell(SrcConstants.DrawingScale, nameof(SrcConstants.DrawingScale));
            this.DrawingScaleType = this.query.AddCell(SrcConstants.DrawingScaleType, nameof(SrcConstants.DrawingScaleType));
            this.DrawingSizeType = this.query.AddCell(SrcConstants.DrawingSizeType, nameof(SrcConstants.DrawingSizeType));
            this.InhibitSnap = this.query.AddCell(SrcConstants.InhibitSnap, nameof(SrcConstants.InhibitSnap));
            this.PageHeight = this.query.AddCell(SrcConstants.PageHeight, nameof(SrcConstants.PageHeight));
            this.PageScale = this.query.AddCell(SrcConstants.PageScale, nameof(SrcConstants.PageScale));
            this.PageWidth = this.query.AddCell(SrcConstants.PageWidth, nameof(SrcConstants.PageWidth));
            this.ShdwObliqueAngle = this.query.AddCell(SrcConstants.ShdwObliqueAngle, nameof(SrcConstants.ShdwObliqueAngle));
            this.ShdwOffsetX = this.query.AddCell(SrcConstants.ShdwOffsetX, nameof(SrcConstants.ShdwOffsetX));
            this.ShdwOffsetY = this.query.AddCell(SrcConstants.ShdwOffsetY, nameof(SrcConstants.ShdwOffsetY));
            this.ShdwScaleFactor = this.query.AddCell(SrcConstants.ShdwScaleFactor, nameof(SrcConstants.ShdwScaleFactor));
            this.ShdwType = this.query.AddCell(SrcConstants.ShdwType, nameof(SrcConstants.ShdwType));
            this.UIVisibility = this.query.AddCell(SrcConstants.UIVisibility, nameof(SrcConstants.UIVisibility));
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
            this.AvenueSizeX = this.query.AddCell(SrcConstants.AvenueSizeX, nameof(SrcConstants.AvenueSizeX));
            this.AvenueSizeY = this.query.AddCell(SrcConstants.AvenueSizeY, nameof(SrcConstants.AvenueSizeY));
            this.BlockSizeX = this.query.AddCell(SrcConstants.BlockSizeX, nameof(SrcConstants.BlockSizeX));
            this.BlockSizeY = this.query.AddCell(SrcConstants.BlockSizeY, nameof(SrcConstants.BlockSizeY));
            this.CtrlAsInput = this.query.AddCell(SrcConstants.CtrlAsInput, nameof(SrcConstants.CtrlAsInput));
            this.DynamicsOff = this.query.AddCell(SrcConstants.DynamicsOff, nameof(SrcConstants.DynamicsOff));
            this.EnableGrid = this.query.AddCell(SrcConstants.EnableGrid, nameof(SrcConstants.EnableGrid));
            this.LineAdjustFrom = this.query.AddCell(SrcConstants.LineAdjustFrom, nameof(SrcConstants.LineAdjustFrom));
            this.LineAdjustTo = this.query.AddCell(SrcConstants.LineAdjustTo, nameof(SrcConstants.LineAdjustTo));
            this.LineJumpCode = this.query.AddCell(SrcConstants.LineJumpCode, nameof(SrcConstants.LineJumpCode));
            this.LineJumpFactorX = this.query.AddCell(SrcConstants.LineJumpFactorX, nameof(SrcConstants.LineJumpFactorX));
            this.LineJumpFactorY = this.query.AddCell(SrcConstants.LineJumpFactorY, nameof(SrcConstants.LineJumpFactorY));
            this.LineJumpStyle = this.query.AddCell(SrcConstants.LineJumpStyle, nameof(SrcConstants.LineJumpStyle));
            this.LineRouteExt = this.query.AddCell(SrcConstants.LineRouteExt, nameof(SrcConstants.LineRouteExt));
            this.LineToLineX = this.query.AddCell(SrcConstants.LineToLineX, nameof(SrcConstants.LineToLineX));
            this.LineToLineY = this.query.AddCell(SrcConstants.LineToLineY, nameof(SrcConstants.LineToLineY));
            this.LineToNodeX = this.query.AddCell(SrcConstants.LineToNodeX, nameof(SrcConstants.LineToNodeX));
            this.LineToNodeY = this.query.AddCell(SrcConstants.LineToNodeY, nameof(SrcConstants.LineToNodeY));
            this.PageLineJumpDirX = this.query.AddCell(SrcConstants.PageLineJumpDirX, nameof(SrcConstants.PageLineJumpDirX));
            this.PageLineJumpDirY = this.query.AddCell(SrcConstants.PageLineJumpDirY, nameof(SrcConstants.PageLineJumpDirY));
            this.PageShapeSplit = this.query.AddCell(SrcConstants.PageShapeSplit, nameof(SrcConstants.PageShapeSplit));
            this.PlaceDepth = this.query.AddCell(SrcConstants.PlaceDepth, nameof(SrcConstants.PlaceDepth));
            this.PlaceFlip = this.query.AddCell(SrcConstants.PlaceFlip, nameof(SrcConstants.PlaceFlip));
            this.PlaceStyle = this.query.AddCell(SrcConstants.PlaceStyle, nameof(SrcConstants.PlaceStyle));
            this.PlowCode = this.query.AddCell(SrcConstants.PlowCode, nameof(SrcConstants.PlowCode));
            this.ResizePage = this.query.AddCell(SrcConstants.ResizePage, nameof(SrcConstants.ResizePage));
            this.RouteStyle = this.query.AddCell(SrcConstants.RouteStyle, nameof(SrcConstants.RouteStyle));
            this.AvoidPageBreaks = this.query.AddCell(SrcConstants.AvoidPageBreaks, nameof(SrcConstants.AvoidPageBreaks));
            this.DrawingResizeType = this.query.AddCell(SrcConstants.DrawingResizeType, nameof(SrcConstants.DrawingResizeType));
        }


        public override Pages.PageCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Pages.PageCells();
            cells.PageLeftMargin = row[this.PageLeftMargin];
            cells.CenterX = row[this.CenterX];
            cells.CenterY = row[this.CenterY];
            cells.OnPage = row[this.OnPage];
            cells.PageBottomMargin = row[this.PageBottomMargin];
            cells.PageRightMargin = row[this.PageRightMargin];
            cells.PagesX = row[this.PagesX];
            cells.PagesY = row[this.PagesY];
            cells.PageTopMargin = row[this.PageTopMargin];
            cells.PaperKind = row[this.PaperKind];
            cells.PrintGrid = row[this.PrintGrid];
            cells.PrintPageOrientation = row[this.PrintPageOrientation];
            cells.ScaleX = row[this.ScaleX];
            cells.ScaleY = row[this.ScaleY];
            cells.PaperSource = row[this.PaperSource];
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
            cells.AvenueSizeX = row[this.AvenueSizeX];
            cells.AvenueSizeY = row[this.AvenueSizeY];
            cells.BlockSizeX = row[this.BlockSizeX];
            cells.BlockSizeY = row[this.BlockSizeY];
            cells.CtrlAsInput = row[this.CtrlAsInput];
            cells.DynamicsOff = row[this.DynamicsOff];
            cells.EnableGrid = row[this.EnableGrid];
            cells.LineAdjustFrom = row[this.LineAdjustFrom];
            cells.LineAdjustTo = row[this.LineAdjustTo];
            cells.LineJumpCode = row[this.LineJumpCode];
            cells.LineJumpFactorX = row[this.LineJumpFactorX];
            cells.LineJumpFactorY = row[this.LineJumpFactorY];
            cells.LineJumpStyle = row[this.LineJumpStyle];
            cells.LineRouteExt = row[this.LineRouteExt];
            cells.LineToLineX = row[this.LineToLineX];
            cells.LineToLineY = row[this.LineToLineY];
            cells.LineToNodeX = row[this.LineToNodeX];
            cells.LineToNodeY = row[this.LineToNodeY];
            cells.PageLineJumpDirX = row[this.PageLineJumpDirX];
            cells.PageLineJumpDirY = row[this.PageLineJumpDirY];
            cells.PageShapeSplit = row[this.PageShapeSplit];
            cells.PlaceDepth = row[this.PlaceDepth];
            cells.PlaceFlip = row[this.PlaceFlip];
            cells.PlaceStyle = row[this.PlaceStyle];
            cells.PlowCode = row[this.PlowCode];
            cells.ResizePage = row[this.ResizePage];
            cells.RouteStyle = row[this.RouteStyle];
            cells.AvoidPageBreaks = row[this.AvoidPageBreaks];
            cells.DrawingResizeType = row[this.DrawingResizeType];
            return cells;
        }
    }
}