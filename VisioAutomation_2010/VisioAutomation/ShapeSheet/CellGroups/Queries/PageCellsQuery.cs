using VisioAutomation.ShapeSheet.Queries.Columns;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using VisioAutomation.Extensions;

namespace VisioAutomation.ShapeSheet.CellGroups.Queries
{
    class PageCellsQuery : CellGroupSingleRowQuery<VisioAutomation.Pages.PageCells, double>
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
            this.PageLeftMargin = this.query.AddCell(SRCCON.PageLeftMargin, nameof(SRCCON.PageLeftMargin));
            this.CenterX = this.query.AddCell(SRCCON.CenterX, nameof(SRCCON.CenterX));
            this.CenterY = this.query.AddCell(SRCCON.CenterY, nameof(SRCCON.CenterY));
            this.OnPage = this.query.AddCell(SRCCON.OnPage, nameof(SRCCON.OnPage));
            this.PageBottomMargin = this.query.AddCell(SRCCON.PageBottomMargin, nameof(SRCCON.PageBottomMargin));
            this.PageRightMargin = this.query.AddCell(SRCCON.PageRightMargin, nameof(SRCCON.PageRightMargin));
            this.PagesX = this.query.AddCell(SRCCON.PagesX, nameof(SRCCON.PagesX));
            this.PagesY = this.query.AddCell(SRCCON.PagesY, nameof(SRCCON.PagesY));
            this.PageTopMargin = this.query.AddCell(SRCCON.PageTopMargin, nameof(SRCCON.PageTopMargin));
            this.PaperKind = this.query.AddCell(SRCCON.PaperKind, nameof(SRCCON.PaperKind));
            this.PrintGrid = this.query.AddCell(SRCCON.PrintGrid, nameof(SRCCON.PrintGrid));
            this.PrintPageOrientation = this.query.AddCell(SRCCON.PrintPageOrientation, nameof(SRCCON.PrintPageOrientation));
            this.ScaleX = this.query.AddCell(SRCCON.ScaleX, nameof(SRCCON.ScaleX));
            this.ScaleY = this.query.AddCell(SRCCON.ScaleY, nameof(SRCCON.ScaleY));
            this.PaperSource = this.query.AddCell(SRCCON.PaperSource, nameof(SRCCON.PaperSource));
            this.DrawingScale = this.query.AddCell(SRCCON.DrawingScale, nameof(SRCCON.DrawingScale));
            this.DrawingScaleType = this.query.AddCell(SRCCON.DrawingScaleType, nameof(SRCCON.DrawingScaleType));
            this.DrawingSizeType = this.query.AddCell(SRCCON.DrawingSizeType, nameof(SRCCON.DrawingSizeType));
            this.InhibitSnap = this.query.AddCell(SRCCON.InhibitSnap, nameof(SRCCON.InhibitSnap));
            this.PageHeight = this.query.AddCell(SRCCON.PageHeight, nameof(SRCCON.PageHeight));
            this.PageScale = this.query.AddCell(SRCCON.PageScale, nameof(SRCCON.PageScale));
            this.PageWidth = this.query.AddCell(SRCCON.PageWidth, nameof(SRCCON.PageWidth));
            this.ShdwObliqueAngle = this.query.AddCell(SRCCON.ShdwObliqueAngle, nameof(SRCCON.ShdwObliqueAngle));
            this.ShdwOffsetX = this.query.AddCell(SRCCON.ShdwOffsetX, nameof(SRCCON.ShdwOffsetX));
            this.ShdwOffsetY = this.query.AddCell(SRCCON.ShdwOffsetY, nameof(SRCCON.ShdwOffsetY));
            this.ShdwScaleFactor = this.query.AddCell(SRCCON.ShdwScaleFactor, nameof(SRCCON.ShdwScaleFactor));
            this.ShdwType = this.query.AddCell(SRCCON.ShdwType, nameof(SRCCON.ShdwType));
            this.UIVisibility = this.query.AddCell(SRCCON.UIVisibility, nameof(SRCCON.UIVisibility));
            this.XGridDensity = this.query.AddCell(SRCCON.XGridDensity, nameof(SRCCON.XGridDensity));
            this.XGridOrigin = this.query.AddCell(SRCCON.XGridOrigin, nameof(SRCCON.XGridOrigin));
            this.XGridSpacing = this.query.AddCell(SRCCON.XGridSpacing, nameof(SRCCON.XGridSpacing));
            this.XRulerDensity = this.query.AddCell(SRCCON.XRulerDensity, nameof(SRCCON.XRulerDensity));
            this.XRulerOrigin = this.query.AddCell(SRCCON.XRulerOrigin, nameof(SRCCON.XRulerOrigin));
            this.YGridDensity = this.query.AddCell(SRCCON.YGridDensity, nameof(SRCCON.YGridDensity));
            this.YGridOrigin = this.query.AddCell(SRCCON.YGridOrigin, nameof(SRCCON.YGridOrigin));
            this.YGridSpacing = this.query.AddCell(SRCCON.YGridSpacing, nameof(SRCCON.YGridSpacing));
            this.YRulerDensity = this.query.AddCell(SRCCON.YRulerDensity, nameof(SRCCON.YRulerDensity));
            this.YRulerOrigin = this.query.AddCell(SRCCON.YRulerOrigin, nameof(SRCCON.YRulerOrigin));
            this.AvenueSizeX = this.query.AddCell(SRCCON.AvenueSizeX, nameof(SRCCON.AvenueSizeX));
            this.AvenueSizeY = this.query.AddCell(SRCCON.AvenueSizeY, nameof(SRCCON.AvenueSizeY));
            this.BlockSizeX = this.query.AddCell(SRCCON.BlockSizeX, nameof(SRCCON.BlockSizeX));
            this.BlockSizeY = this.query.AddCell(SRCCON.BlockSizeY, nameof(SRCCON.BlockSizeY));
            this.CtrlAsInput = this.query.AddCell(SRCCON.CtrlAsInput, nameof(SRCCON.CtrlAsInput));
            this.DynamicsOff = this.query.AddCell(SRCCON.DynamicsOff, nameof(SRCCON.DynamicsOff));
            this.EnableGrid = this.query.AddCell(SRCCON.EnableGrid, nameof(SRCCON.EnableGrid));
            this.LineAdjustFrom = this.query.AddCell(SRCCON.LineAdjustFrom, nameof(SRCCON.LineAdjustFrom));
            this.LineAdjustTo = this.query.AddCell(SRCCON.LineAdjustTo, nameof(SRCCON.LineAdjustTo));
            this.LineJumpCode = this.query.AddCell(SRCCON.LineJumpCode, nameof(SRCCON.LineJumpCode));
            this.LineJumpFactorX = this.query.AddCell(SRCCON.LineJumpFactorX, nameof(SRCCON.LineJumpFactorX));
            this.LineJumpFactorY = this.query.AddCell(SRCCON.LineJumpFactorY, nameof(SRCCON.LineJumpFactorY));
            this.LineJumpStyle = this.query.AddCell(SRCCON.LineJumpStyle, nameof(SRCCON.LineJumpStyle));
            this.LineRouteExt = this.query.AddCell(SRCCON.LineRouteExt, nameof(SRCCON.LineRouteExt));
            this.LineToLineX = this.query.AddCell(SRCCON.LineToLineX, nameof(SRCCON.LineToLineX));
            this.LineToLineY = this.query.AddCell(SRCCON.LineToLineY, nameof(SRCCON.LineToLineY));
            this.LineToNodeX = this.query.AddCell(SRCCON.LineToNodeX, nameof(SRCCON.LineToNodeX));
            this.LineToNodeY = this.query.AddCell(SRCCON.LineToNodeY, nameof(SRCCON.LineToNodeY));
            this.PageLineJumpDirX = this.query.AddCell(SRCCON.PageLineJumpDirX, nameof(SRCCON.PageLineJumpDirX));
            this.PageLineJumpDirY = this.query.AddCell(SRCCON.PageLineJumpDirY, nameof(SRCCON.PageLineJumpDirY));
            this.PageShapeSplit = this.query.AddCell(SRCCON.PageShapeSplit, nameof(SRCCON.PageShapeSplit));
            this.PlaceDepth = this.query.AddCell(SRCCON.PlaceDepth, nameof(SRCCON.PlaceDepth));
            this.PlaceFlip = this.query.AddCell(SRCCON.PlaceFlip, nameof(SRCCON.PlaceFlip));
            this.PlaceStyle = this.query.AddCell(SRCCON.PlaceStyle, nameof(SRCCON.PlaceStyle));
            this.PlowCode = this.query.AddCell(SRCCON.PlowCode, nameof(SRCCON.PlowCode));
            this.ResizePage = this.query.AddCell(SRCCON.ResizePage, nameof(SRCCON.ResizePage));
            this.RouteStyle = this.query.AddCell(SRCCON.RouteStyle, nameof(SRCCON.RouteStyle));
            this.AvoidPageBreaks = this.query.AddCell(SRCCON.AvoidPageBreaks, nameof(SRCCON.AvoidPageBreaks));
            this.DrawingResizeType = this.query.AddCell(SRCCON.DrawingResizeType, nameof(SRCCON.DrawingResizeType));
        }


        public override Pages.PageCells CellDataToCellGroup(ShapeSheet.CellData<double>[] row)
        {
            var cells = new Pages.PageCells();
            cells.PageLeftMargin = row[this.PageLeftMargin];
            cells.CenterX = row[this.CenterX];
            cells.CenterY = row[this.CenterY];
            cells.OnPage = row[this.OnPage].ToInt();
            cells.PageBottomMargin = row[this.PageBottomMargin];
            cells.PageRightMargin = row[this.PageRightMargin];
            cells.PagesX = row[this.PagesX];
            cells.PagesY = row[this.PagesY];
            cells.PageTopMargin = row[this.PageTopMargin];
            cells.PaperKind = row[this.PaperKind].ToInt();
            cells.PrintGrid = row[this.PrintGrid].ToInt();
            cells.PrintPageOrientation = row[this.PrintPageOrientation].ToInt();
            cells.ScaleX = row[this.ScaleX];
            cells.ScaleY = row[this.ScaleY];
            cells.PaperSource = row[this.PaperSource].ToInt();
            cells.DrawingScale = row[this.DrawingScale];
            cells.DrawingScaleType = row[this.DrawingScaleType].ToInt();
            cells.DrawingSizeType = row[this.DrawingSizeType].ToInt();
            cells.InhibitSnap = row[this.InhibitSnap].ToInt();
            cells.PageHeight = row[this.PageHeight];
            cells.PageScale = row[this.PageScale];
            cells.PageWidth = row[this.PageWidth];
            cells.ShdwObliqueAngle = row[this.ShdwObliqueAngle];
            cells.ShdwOffsetX = row[this.ShdwOffsetX];
            cells.ShdwOffsetY = row[this.ShdwOffsetY];
            cells.ShdwScaleFactor = row[this.ShdwScaleFactor];
            cells.ShdwType = row[this.ShdwType].ToInt();
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
            cells.CtrlAsInput = row[this.CtrlAsInput].ToInt();
            cells.DynamicsOff = row[this.DynamicsOff].ToInt();
            cells.EnableGrid = row[this.EnableGrid].ToInt();
            cells.LineAdjustFrom = row[this.LineAdjustFrom].ToInt();
            cells.LineAdjustTo = row[this.LineAdjustTo];
            cells.LineJumpCode = row[this.LineJumpCode];
            cells.LineJumpFactorX = row[this.LineJumpFactorX];
            cells.LineJumpFactorY = row[this.LineJumpFactorY];
            cells.LineJumpStyle = row[this.LineJumpStyle].ToInt();
            cells.LineRouteExt = row[this.LineRouteExt];
            cells.LineToLineX = row[this.LineToLineX];
            cells.LineToLineY = row[this.LineToLineY];
            cells.LineToNodeX = row[this.LineToNodeX];
            cells.LineToNodeY = row[this.LineToNodeY];
            cells.PageLineJumpDirX = row[this.PageLineJumpDirX];
            cells.PageLineJumpDirY = row[this.PageLineJumpDirY];
            cells.PageShapeSplit = row[this.PageShapeSplit].ToInt();
            cells.PlaceDepth = row[this.PlaceDepth].ToInt();
            cells.PlaceFlip = row[this.PlaceFlip].ToInt();
            cells.PlaceStyle = row[this.PlaceStyle].ToInt();
            cells.PlowCode = row[this.PlowCode].ToInt();
            cells.ResizePage = row[this.ResizePage].ToInt();
            cells.RouteStyle = row[this.RouteStyle].ToInt();
            cells.AvoidPageBreaks = row[this.AvoidPageBreaks].ToInt();
            cells.DrawingResizeType = row[this.DrawingResizeType].ToInt();
            return cells;
        }
    }
}