using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.ShapeSheetQuery.Common
{
    class PageCellsQuery : CellQuery
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

        public PageCellsQuery()
        {
            this.PageLeftMargin = this.AddCell(SRCCON.PageLeftMargin, nameof(SRCCON.PageLeftMargin));
            this.CenterX = this.AddCell(SRCCON.CenterX, nameof(SRCCON.CenterX));
            this.CenterY = this.AddCell(SRCCON.CenterY, nameof(SRCCON.CenterY));
            this.OnPage = this.AddCell(SRCCON.OnPage, nameof(SRCCON.OnPage));
            this.PageBottomMargin = this.AddCell(SRCCON.PageBottomMargin, nameof(SRCCON.PageBottomMargin));
            this.PageRightMargin = this.AddCell(SRCCON.PageRightMargin, nameof(SRCCON.PageRightMargin));
            this.PagesX = this.AddCell(SRCCON.PagesX, nameof(SRCCON.PagesX));
            this.PagesY = this.AddCell(SRCCON.PagesY, nameof(SRCCON.PagesY));
            this.PageTopMargin = this.AddCell(SRCCON.PageTopMargin, nameof(SRCCON.PageTopMargin));
            this.PaperKind = this.AddCell(SRCCON.PaperKind, nameof(SRCCON.PaperKind));
            this.PrintGrid = this.AddCell(SRCCON.PrintGrid, nameof(SRCCON.PrintGrid));
            this.PrintPageOrientation = this.AddCell(SRCCON.PrintPageOrientation, nameof(SRCCON.PrintPageOrientation));
            this.ScaleX = this.AddCell(SRCCON.ScaleX, nameof(SRCCON.ScaleX));
            this.ScaleY = this.AddCell(SRCCON.ScaleY, nameof(SRCCON.ScaleY));
            this.PaperSource = this.AddCell(SRCCON.PaperSource, nameof(SRCCON.PaperSource));
            this.DrawingScale = this.AddCell(SRCCON.DrawingScale, nameof(SRCCON.DrawingScale));
            this.DrawingScaleType = this.AddCell(SRCCON.DrawingScaleType, nameof(SRCCON.DrawingScaleType));
            this.DrawingSizeType = this.AddCell(SRCCON.DrawingSizeType, nameof(SRCCON.DrawingSizeType));
            this.InhibitSnap = this.AddCell(SRCCON.InhibitSnap, nameof(SRCCON.InhibitSnap));
            this.PageHeight = this.AddCell(SRCCON.PageHeight, nameof(SRCCON.PageHeight));
            this.PageScale = this.AddCell(SRCCON.PageScale, nameof(SRCCON.PageScale));
            this.PageWidth = this.AddCell(SRCCON.PageWidth, nameof(SRCCON.PageWidth));
            this.ShdwObliqueAngle = this.AddCell(SRCCON.ShdwObliqueAngle, nameof(SRCCON.ShdwObliqueAngle));
            this.ShdwOffsetX = this.AddCell(SRCCON.ShdwOffsetX, nameof(SRCCON.ShdwOffsetX));
            this.ShdwOffsetY = this.AddCell(SRCCON.ShdwOffsetY, nameof(SRCCON.ShdwOffsetY));
            this.ShdwScaleFactor = this.AddCell(SRCCON.ShdwScaleFactor, nameof(SRCCON.ShdwScaleFactor));
            this.ShdwType = this.AddCell(SRCCON.ShdwType, nameof(SRCCON.ShdwType));
            this.UIVisibility = this.AddCell(SRCCON.UIVisibility, nameof(SRCCON.UIVisibility));
            this.XGridDensity = this.AddCell(SRCCON.XGridDensity, nameof(SRCCON.XGridDensity));
            this.XGridOrigin = this.AddCell(SRCCON.XGridOrigin, nameof(SRCCON.XGridOrigin));
            this.XGridSpacing = this.AddCell(SRCCON.XGridSpacing, nameof(SRCCON.XGridSpacing));
            this.XRulerDensity = this.AddCell(SRCCON.XRulerDensity, nameof(SRCCON.XRulerDensity));
            this.XRulerOrigin = this.AddCell(SRCCON.XRulerOrigin, nameof(SRCCON.XRulerOrigin));
            this.YGridDensity = this.AddCell(SRCCON.YGridDensity, nameof(SRCCON.YGridDensity));
            this.YGridOrigin = this.AddCell(SRCCON.YGridOrigin, nameof(SRCCON.YGridOrigin));
            this.YGridSpacing = this.AddCell(SRCCON.YGridSpacing, nameof(SRCCON.YGridSpacing));
            this.YRulerDensity = this.AddCell(SRCCON.YRulerDensity, nameof(SRCCON.YRulerDensity));
            this.YRulerOrigin = this.AddCell(SRCCON.YRulerOrigin, nameof(SRCCON.YRulerOrigin));
            this.AvenueSizeX = this.AddCell(SRCCON.AvenueSizeX, nameof(SRCCON.AvenueSizeX));
            this.AvenueSizeY = this.AddCell(SRCCON.AvenueSizeY, nameof(SRCCON.AvenueSizeY));
            this.BlockSizeX = this.AddCell(SRCCON.BlockSizeX, nameof(SRCCON.BlockSizeX));
            this.BlockSizeY = this.AddCell(SRCCON.BlockSizeY, nameof(SRCCON.BlockSizeY));
            this.CtrlAsInput = this.AddCell(SRCCON.CtrlAsInput, nameof(SRCCON.CtrlAsInput));
            this.DynamicsOff = this.AddCell(SRCCON.DynamicsOff, nameof(SRCCON.DynamicsOff));
            this.EnableGrid = this.AddCell(SRCCON.EnableGrid, nameof(SRCCON.EnableGrid));
            this.LineAdjustFrom = this.AddCell(SRCCON.LineAdjustFrom, nameof(SRCCON.LineAdjustFrom));
            this.LineAdjustTo = this.AddCell(SRCCON.LineAdjustTo, nameof(SRCCON.LineAdjustTo));
            this.LineJumpCode = this.AddCell(SRCCON.LineJumpCode, nameof(SRCCON.LineJumpCode));
            this.LineJumpFactorX = this.AddCell(SRCCON.LineJumpFactorX, nameof(SRCCON.LineJumpFactorX));
            this.LineJumpFactorY = this.AddCell(SRCCON.LineJumpFactorY, nameof(SRCCON.LineJumpFactorY));
            this.LineJumpStyle = this.AddCell(SRCCON.LineJumpStyle, nameof(SRCCON.LineJumpStyle));
            this.LineRouteExt = this.AddCell(SRCCON.LineRouteExt, nameof(SRCCON.LineRouteExt));
            this.LineToLineX = this.AddCell(SRCCON.LineToLineX, nameof(SRCCON.LineToLineX));
            this.LineToLineY = this.AddCell(SRCCON.LineToLineY, nameof(SRCCON.LineToLineY));
            this.LineToNodeX = this.AddCell(SRCCON.LineToNodeX, nameof(SRCCON.LineToNodeX));
            this.LineToNodeY = this.AddCell(SRCCON.LineToNodeY, nameof(SRCCON.LineToNodeY));
            this.PageLineJumpDirX = this.AddCell(SRCCON.PageLineJumpDirX, nameof(SRCCON.PageLineJumpDirX));
            this.PageLineJumpDirY = this.AddCell(SRCCON.PageLineJumpDirY, nameof(SRCCON.PageLineJumpDirY));
            this.PageShapeSplit = this.AddCell(SRCCON.PageShapeSplit, nameof(SRCCON.PageShapeSplit));
            this.PlaceDepth = this.AddCell(SRCCON.PlaceDepth, nameof(SRCCON.PlaceDepth));
            this.PlaceFlip = this.AddCell(SRCCON.PlaceFlip, nameof(SRCCON.PlaceFlip));
            this.PlaceStyle = this.AddCell(SRCCON.PlaceStyle, nameof(SRCCON.PlaceStyle));
            this.PlowCode = this.AddCell(SRCCON.PlowCode, nameof(SRCCON.PlowCode));
            this.ResizePage = this.AddCell(SRCCON.ResizePage, nameof(SRCCON.ResizePage));
            this.RouteStyle = this.AddCell(SRCCON.RouteStyle, nameof(SRCCON.RouteStyle));
            this.AvoidPageBreaks = this.AddCell(SRCCON.AvoidPageBreaks, nameof(SRCCON.AvoidPageBreaks));
            this.DrawingResizeType = this.AddCell(SRCCON.DrawingResizeType, nameof(SRCCON.DrawingResizeType));
        }


        public Pages.PageCells GetCells(SectionResultRow<ShapeSheet.CellData<double>> row)
        {

            var cells = new Pages.PageCells();
            cells.PageLeftMargin = row.Cells[this.PageLeftMargin];
            cells.CenterX = row.Cells[this.CenterX];
            cells.CenterY = row.Cells[this.CenterY];
            cells.OnPage = Extensions.CellDataMethods.ToInt(row.Cells[this.OnPage]);
            cells.PageBottomMargin = row.Cells[this.PageBottomMargin];
            cells.PageRightMargin = row.Cells[this.PageRightMargin];
            cells.PagesX = row.Cells[this.PagesX];
            cells.PagesY = row.Cells[this.PagesY];
            cells.PageTopMargin = row.Cells[this.PageTopMargin];
            cells.PaperKind = Extensions.CellDataMethods.ToInt(row.Cells[this.PaperKind]);
            cells.PrintGrid = Extensions.CellDataMethods.ToInt(row.Cells[this.PrintGrid]);
            cells.PrintPageOrientation = Extensions.CellDataMethods.ToInt(row.Cells[this.PrintPageOrientation]);
            cells.ScaleX = row.Cells[this.ScaleX];
            cells.ScaleY = row.Cells[this.ScaleY];
            cells.PaperSource = Extensions.CellDataMethods.ToInt(row.Cells[this.PaperSource]);
            cells.DrawingScale = row.Cells[this.DrawingScale];
            cells.DrawingScaleType = Extensions.CellDataMethods.ToInt(row.Cells[this.DrawingScaleType]);
            cells.DrawingSizeType = Extensions.CellDataMethods.ToInt(row.Cells[this.DrawingSizeType]);
            cells.InhibitSnap = Extensions.CellDataMethods.ToInt(row.Cells[this.InhibitSnap]);
            cells.PageHeight = row.Cells[this.PageHeight];
            cells.PageScale = row.Cells[this.PageScale];
            cells.PageWidth = row.Cells[this.PageWidth];
            cells.ShdwObliqueAngle = row.Cells[this.ShdwObliqueAngle];
            cells.ShdwOffsetX = row.Cells[this.ShdwOffsetX];
            cells.ShdwOffsetY = row.Cells[this.ShdwOffsetY];
            cells.ShdwScaleFactor = row.Cells[this.ShdwScaleFactor];
            cells.ShdwType = Extensions.CellDataMethods.ToInt(row.Cells[this.ShdwType]);
            cells.UIVisibility = row.Cells[this.UIVisibility];
            cells.XGridDensity = row.Cells[this.XGridDensity];
            cells.XGridOrigin = row.Cells[this.XGridOrigin];
            cells.XGridSpacing = row.Cells[this.XGridSpacing];
            cells.XRulerDensity = row.Cells[this.XRulerDensity];
            cells.XRulerOrigin = row.Cells[this.XRulerOrigin];
            cells.YGridDensity = row.Cells[this.YGridDensity];
            cells.YGridOrigin = row.Cells[this.YGridOrigin];
            cells.YGridSpacing = row.Cells[this.YGridSpacing];
            cells.YRulerDensity = row.Cells[this.YRulerDensity];
            cells.YRulerOrigin = row.Cells[this.YRulerOrigin];
            cells.AvenueSizeX = row.Cells[this.AvenueSizeX];
            cells.AvenueSizeY = row.Cells[this.AvenueSizeY];
            cells.BlockSizeX = row.Cells[this.BlockSizeX];
            cells.BlockSizeY = row.Cells[this.BlockSizeY];
            cells.CtrlAsInput = Extensions.CellDataMethods.ToInt(row.Cells[this.CtrlAsInput]);
            cells.DynamicsOff = Extensions.CellDataMethods.ToInt(row.Cells[this.DynamicsOff]);
            cells.EnableGrid = Extensions.CellDataMethods.ToInt(row.Cells[this.EnableGrid]);
            cells.LineAdjustFrom = Extensions.CellDataMethods.ToInt(row.Cells[this.LineAdjustFrom]);
            cells.LineAdjustTo = row.Cells[this.LineAdjustTo];
            cells.LineJumpCode = row.Cells[this.LineJumpCode];
            cells.LineJumpFactorX = row.Cells[this.LineJumpFactorX];
            cells.LineJumpFactorY = row.Cells[this.LineJumpFactorY];
            cells.LineJumpStyle = Extensions.CellDataMethods.ToInt(row.Cells[this.LineJumpStyle]);
            cells.LineRouteExt = row.Cells[this.LineRouteExt];
            cells.LineToLineX = row.Cells[this.LineToLineX];
            cells.LineToLineY = row.Cells[this.LineToLineY];
            cells.LineToNodeX = row.Cells[this.LineToNodeX];
            cells.LineToNodeY = row.Cells[this.LineToNodeY];
            cells.PageLineJumpDirX = row.Cells[this.PageLineJumpDirX];
            cells.PageLineJumpDirY = row.Cells[this.PageLineJumpDirY];
            cells.PageShapeSplit = Extensions.CellDataMethods.ToInt(row.Cells[this.PageShapeSplit]);
            cells.PlaceDepth = Extensions.CellDataMethods.ToInt(row.Cells[this.PlaceDepth]);
            cells.PlaceFlip = Extensions.CellDataMethods.ToInt(row.Cells[this.PlaceFlip]);
            cells.PlaceStyle = Extensions.CellDataMethods.ToInt(row.Cells[this.PlaceStyle]);
            cells.PlowCode = Extensions.CellDataMethods.ToInt(row.Cells[this.PlowCode]);
            cells.ResizePage = Extensions.CellDataMethods.ToInt(row.Cells[this.ResizePage]);
            cells.RouteStyle = Extensions.CellDataMethods.ToInt(row.Cells[this.RouteStyle]);
            cells.AvoidPageBreaks = Extensions.CellDataMethods.ToInt(row.Cells[this.AvoidPageBreaks]);
            cells.DrawingResizeType = Extensions.CellDataMethods.ToInt(row.Cells[this.DrawingResizeType]);
            return cells;
        }

    }
}