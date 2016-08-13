using VisioAutomation.ShapeSheetQuery.Columns;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.ShapeSheetQuery.CommonQueries
{
    class PageCellsQuery : Query
    {
        public ColumnQuery PageLeftMargin { get; set; }
        public ColumnQuery CenterX { get; set; }
        public ColumnQuery CenterY { get; set; }
        public ColumnQuery OnPage { get; set; }
        public ColumnQuery PageBottomMargin { get; set; }
        public ColumnQuery PageRightMargin { get; set; }
        public ColumnQuery PagesX { get; set; }
        public ColumnQuery PagesY { get; set; }
        public ColumnQuery PageTopMargin { get; set; }
        public ColumnQuery PaperKind { get; set; }
        public ColumnQuery PrintGrid { get; set; }
        public ColumnQuery PrintPageOrientation { get; set; }
        public ColumnQuery ScaleX { get; set; }
        public ColumnQuery ScaleY { get; set; }
        public ColumnQuery PaperSource { get; set; }
        public ColumnQuery DrawingScale { get; set; }
        public ColumnQuery DrawingScaleType { get; set; }
        public ColumnQuery DrawingSizeType { get; set; }
        public ColumnQuery InhibitSnap { get; set; }
        public ColumnQuery PageHeight { get; set; }
        public ColumnQuery PageScale { get; set; }
        public ColumnQuery PageWidth { get; set; }
        public ColumnQuery ShdwObliqueAngle { get; set; }
        public ColumnQuery ShdwOffsetX { get; set; }
        public ColumnQuery ShdwOffsetY { get; set; }
        public ColumnQuery ShdwScaleFactor { get; set; }
        public ColumnQuery ShdwType { get; set; }
        public ColumnQuery UIVisibility { get; set; }
        public ColumnQuery XGridDensity { get; set; }
        public ColumnQuery XGridOrigin { get; set; }
        public ColumnQuery XGridSpacing { get; set; }
        public ColumnQuery XRulerDensity { get; set; }
        public ColumnQuery XRulerOrigin { get; set; }
        public ColumnQuery YGridDensity { get; set; }
        public ColumnQuery YGridOrigin { get; set; }
        public ColumnQuery YGridSpacing { get; set; }
        public ColumnQuery YRulerDensity { get; set; }
        public ColumnQuery YRulerOrigin { get; set; }
        public ColumnQuery AvenueSizeX { get; set; }
        public ColumnQuery AvenueSizeY { get; set; }
        public ColumnQuery BlockSizeX { get; set; }
        public ColumnQuery BlockSizeY { get; set; }
        public ColumnQuery CtrlAsInput { get; set; }
        public ColumnQuery DynamicsOff { get; set; }
        public ColumnQuery EnableGrid { get; set; }
        public ColumnQuery LineAdjustFrom { get; set; }
        public ColumnQuery LineAdjustTo { get; set; }
        public ColumnQuery LineJumpCode { get; set; }
        public ColumnQuery LineJumpFactorX { get; set; }
        public ColumnQuery LineJumpFactorY { get; set; }
        public ColumnQuery LineJumpStyle { get; set; }
        public ColumnQuery LineRouteExt { get; set; }
        public ColumnQuery LineToLineX { get; set; }
        public ColumnQuery LineToLineY { get; set; }
        public ColumnQuery LineToNodeX { get; set; }
        public ColumnQuery LineToNodeY { get; set; }
        public ColumnQuery PageLineJumpDirX { get; set; }
        public ColumnQuery PageLineJumpDirY { get; set; }
        public ColumnQuery PageShapeSplit { get; set; }
        public ColumnQuery PlaceDepth { get; set; }
        public ColumnQuery PlaceFlip { get; set; }
        public ColumnQuery PlaceStyle { get; set; }
        public ColumnQuery PlowCode { get; set; }
        public ColumnQuery ResizePage { get; set; }
        public ColumnQuery RouteStyle { get; set; }
        public ColumnQuery AvoidPageBreaks { get; set; }
        public ColumnQuery DrawingResizeType { get; set; }

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


        public Pages.PageCells GetCells(ShapeSheet.CellData<double>[] row)
        {

            var cells = new Pages.PageCells();
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