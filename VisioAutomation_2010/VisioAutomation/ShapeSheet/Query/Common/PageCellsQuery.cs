using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.ShapeSheet.Query.Common
{
    class PageCellsQuery : CellQuery
    {
        public Query.CellColumn PageLeftMargin { get; set; }
        public Query.CellColumn CenterX { get; set; }
        public Query.CellColumn CenterY { get; set; }
        public Query.CellColumn OnPage { get; set; }
        public Query.CellColumn PageBottomMargin { get; set; }
        public Query.CellColumn PageRightMargin { get; set; }
        public Query.CellColumn PagesX { get; set; }
        public Query.CellColumn PagesY { get; set; }
        public Query.CellColumn PageTopMargin { get; set; }
        public Query.CellColumn PaperKind { get; set; }
        public Query.CellColumn PrintGrid { get; set; }
        public Query.CellColumn PrintPageOrientation { get; set; }
        public Query.CellColumn ScaleX { get; set; }
        public Query.CellColumn ScaleY { get; set; }
        public Query.CellColumn PaperSource { get; set; }
        public Query.CellColumn DrawingScale { get; set; }
        public Query.CellColumn DrawingScaleType { get; set; }
        public Query.CellColumn DrawingSizeType { get; set; }
        public Query.CellColumn InhibitSnap { get; set; }
        public Query.CellColumn PageHeight { get; set; }
        public Query.CellColumn PageScale { get; set; }
        public Query.CellColumn PageWidth { get; set; }
        public Query.CellColumn ShdwObliqueAngle { get; set; }
        public Query.CellColumn ShdwOffsetX { get; set; }
        public Query.CellColumn ShdwOffsetY { get; set; }
        public Query.CellColumn ShdwScaleFactor { get; set; }
        public Query.CellColumn ShdwType { get; set; }
        public Query.CellColumn UIVisibility { get; set; }
        public Query.CellColumn XGridDensity { get; set; }
        public Query.CellColumn XGridOrigin { get; set; }
        public Query.CellColumn XGridSpacing { get; set; }
        public Query.CellColumn XRulerDensity { get; set; }
        public Query.CellColumn XRulerOrigin { get; set; }
        public Query.CellColumn YGridDensity { get; set; }
        public Query.CellColumn YGridOrigin { get; set; }
        public Query.CellColumn YGridSpacing { get; set; }
        public Query.CellColumn YRulerDensity { get; set; }
        public Query.CellColumn YRulerOrigin { get; set; }
        public Query.CellColumn AvenueSizeX { get; set; }
        public Query.CellColumn AvenueSizeY { get; set; }
        public Query.CellColumn BlockSizeX { get; set; }
        public Query.CellColumn BlockSizeY { get; set; }
        public Query.CellColumn CtrlAsInput { get; set; }
        public Query.CellColumn DynamicsOff { get; set; }
        public Query.CellColumn EnableGrid { get; set; }
        public Query.CellColumn LineAdjustFrom { get; set; }
        public Query.CellColumn LineAdjustTo { get; set; }
        public Query.CellColumn LineJumpCode { get; set; }
        public Query.CellColumn LineJumpFactorX { get; set; }
        public Query.CellColumn LineJumpFactorY { get; set; }
        public Query.CellColumn LineJumpStyle { get; set; }
        public Query.CellColumn LineRouteExt { get; set; }
        public Query.CellColumn LineToLineX { get; set; }
        public Query.CellColumn LineToLineY { get; set; }
        public Query.CellColumn LineToNodeX { get; set; }
        public Query.CellColumn LineToNodeY { get; set; }
        public Query.CellColumn PageLineJumpDirX { get; set; }
        public Query.CellColumn PageLineJumpDirY { get; set; }
        public Query.CellColumn PageShapeSplit { get; set; }
        public Query.CellColumn PlaceDepth { get; set; }
        public Query.CellColumn PlaceFlip { get; set; }
        public Query.CellColumn PlaceStyle { get; set; }
        public Query.CellColumn PlowCode { get; set; }
        public Query.CellColumn ResizePage { get; set; }
        public Query.CellColumn RouteStyle { get; set; }
        public Query.CellColumn AvoidPageBreaks { get; set; }
        public Query.CellColumn DrawingResizeType { get; set; }

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


        public VisioAutomation.Pages.PageCells GetCells(System.Collections.Generic.IList<ShapeSheet.CellData<double>> row)
        {

            var cells = new VisioAutomation.Pages.PageCells();
            cells.PageLeftMargin = row[this.PageLeftMargin];
            cells.CenterX = row[this.CenterX];
            cells.CenterY = row[this.CenterY];
            cells.OnPage = Extensions.CellDataMethods.ToInt(row[this.OnPage]);
            cells.PageBottomMargin = row[this.PageBottomMargin];
            cells.PageRightMargin = row[this.PageRightMargin];
            cells.PagesX = row[this.PagesX];
            cells.PagesY = row[this.PagesY];
            cells.PageTopMargin = row[this.PageTopMargin];
            cells.PaperKind = Extensions.CellDataMethods.ToInt(row[this.PaperKind]);
            cells.PrintGrid = Extensions.CellDataMethods.ToInt(row[this.PrintGrid]);
            cells.PrintPageOrientation = Extensions.CellDataMethods.ToInt(row[this.PrintPageOrientation]);
            cells.ScaleX = row[this.ScaleX];
            cells.ScaleY = row[this.ScaleY];
            cells.PaperSource = Extensions.CellDataMethods.ToInt(row[this.PaperSource]);
            cells.DrawingScale = row[this.DrawingScale];
            cells.DrawingScaleType = Extensions.CellDataMethods.ToInt(row[this.DrawingScaleType]);
            cells.DrawingSizeType = Extensions.CellDataMethods.ToInt(row[this.DrawingSizeType]);
            cells.InhibitSnap = Extensions.CellDataMethods.ToInt(row[this.InhibitSnap]);
            cells.PageHeight = row[this.PageHeight];
            cells.PageScale = row[this.PageScale];
            cells.PageWidth = row[this.PageWidth];
            cells.ShdwObliqueAngle = row[this.ShdwObliqueAngle];
            cells.ShdwOffsetX = row[this.ShdwOffsetX];
            cells.ShdwOffsetY = row[this.ShdwOffsetY];
            cells.ShdwScaleFactor = row[this.ShdwScaleFactor];
            cells.ShdwType = Extensions.CellDataMethods.ToInt(row[this.ShdwType]);
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
            cells.CtrlAsInput = Extensions.CellDataMethods.ToInt(row[this.CtrlAsInput]);
            cells.DynamicsOff = Extensions.CellDataMethods.ToInt(row[this.DynamicsOff]);
            cells.EnableGrid = Extensions.CellDataMethods.ToInt(row[this.EnableGrid]);
            cells.LineAdjustFrom = Extensions.CellDataMethods.ToInt(row[this.LineAdjustFrom]);
            cells.LineAdjustTo = row[this.LineAdjustTo];
            cells.LineJumpCode = row[this.LineJumpCode];
            cells.LineJumpFactorX = row[this.LineJumpFactorX];
            cells.LineJumpFactorY = row[this.LineJumpFactorY];
            cells.LineJumpStyle = Extensions.CellDataMethods.ToInt(row[this.LineJumpStyle]);
            cells.LineRouteExt = row[this.LineRouteExt];
            cells.LineToLineX = row[this.LineToLineX];
            cells.LineToLineY = row[this.LineToLineY];
            cells.LineToNodeX = row[this.LineToNodeX];
            cells.LineToNodeY = row[this.LineToNodeY];
            cells.PageLineJumpDirX = row[this.PageLineJumpDirX];
            cells.PageLineJumpDirY = row[this.PageLineJumpDirY];
            cells.PageShapeSplit = Extensions.CellDataMethods.ToInt(row[this.PageShapeSplit]);
            cells.PlaceDepth = Extensions.CellDataMethods.ToInt(row[this.PlaceDepth]);
            cells.PlaceFlip = Extensions.CellDataMethods.ToInt(row[this.PlaceFlip]);
            cells.PlaceStyle = Extensions.CellDataMethods.ToInt(row[this.PlaceStyle]);
            cells.PlowCode = Extensions.CellDataMethods.ToInt(row[this.PlowCode]);
            cells.ResizePage = Extensions.CellDataMethods.ToInt(row[this.ResizePage]);
            cells.RouteStyle = Extensions.CellDataMethods.ToInt(row[this.RouteStyle]);
            cells.AvoidPageBreaks = Extensions.CellDataMethods.ToInt(row[this.AvoidPageBreaks]);
            cells.DrawingResizeType = Extensions.CellDataMethods.ToInt(row[this.DrawingResizeType]);
            return cells;
        }

    }
}