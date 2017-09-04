using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Pages
{
    class PageLayoutCellsReader : ReaderSingleRow<VisioAutomation.Pages.PageLayoutCells>
    {
        public CellColumn AvenueSizeX { get; set; }
        public CellColumn AvenueSizeY { get; set; }
        public CellColumn BlockSizeX { get; set; }
        public CellColumn BlockSizeY { get; set; }
        public CellColumn ControlAsInput { get; set; }
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
        public CellColumn LineJumpDirX { get; set; }
        public CellColumn LineJumpDirY { get; set; }
        public CellColumn ShapeSplit { get; set; }
        public CellColumn PlaceDepth { get; set; }
        public CellColumn PlaceFlip { get; set; }
        public CellColumn PlaceStyle { get; set; }
        public CellColumn PlowCode { get; set; }
        public CellColumn ResizePage { get; set; }
        public CellColumn RouteStyle { get; set; }
        public CellColumn AvoidPageBreaks { get; set; }

        public PageLayoutCellsReader()
        {
            this.AvenueSizeX = this.query.AddColumn(SrcConstants.PageLayoutAvenueSizeX, nameof(SrcConstants.PageLayoutAvenueSizeX));
            this.AvenueSizeY = this.query.AddColumn(SrcConstants.PageLayoutAvenueSizeY, nameof(SrcConstants.PageLayoutAvenueSizeY));
            this.BlockSizeX = this.query.AddColumn(SrcConstants.PageLayoutBlockSizeX, nameof(SrcConstants.PageLayoutBlockSizeX));
            this.BlockSizeY = this.query.AddColumn(SrcConstants.PageLayoutBlockSizeY, nameof(SrcConstants.PageLayoutBlockSizeY));
            this.ControlAsInput = this.query.AddColumn(SrcConstants.PageLayoutControlAsInput, nameof(SrcConstants.PageLayoutControlAsInput));
            this.DynamicsOff = this.query.AddColumn(SrcConstants.PageLayoutDynamicsOff, nameof(SrcConstants.PageLayoutDynamicsOff));
            this.EnableGrid = this.query.AddColumn(SrcConstants.PageLayoutEnableGrid, nameof(SrcConstants.PageLayoutEnableGrid));
            this.LineAdjustFrom = this.query.AddColumn(SrcConstants.PageLayoutLineAdjustFrom, nameof(SrcConstants.PageLayoutLineAdjustFrom));
            this.LineAdjustTo = this.query.AddColumn(SrcConstants.PageLayoutLineAdjustTo, nameof(SrcConstants.PageLayoutLineAdjustTo));
            this.LineJumpCode = this.query.AddColumn(SrcConstants.PageLayoutLineJumpCode, nameof(SrcConstants.PageLayoutLineJumpCode));
            this.LineJumpFactorX = this.query.AddColumn(SrcConstants.PageLayoutLineJumpFactorX, nameof(SrcConstants.PageLayoutLineJumpFactorX));
            this.LineJumpFactorY = this.query.AddColumn(SrcConstants.PageLayoutLineJumpFactorY, nameof(SrcConstants.PageLayoutLineJumpFactorY));
            this.LineJumpStyle = this.query.AddColumn(SrcConstants.PageLayoutLineJumpStyle, nameof(SrcConstants.PageLayoutLineJumpStyle));
            this.LineRouteExt = this.query.AddColumn(SrcConstants.PageLayoutLineRouteExt, nameof(SrcConstants.PageLayoutLineRouteExt));
            this.LineToLineX = this.query.AddColumn(SrcConstants.PageLayoutLineToLineX, nameof(SrcConstants.PageLayoutLineToLineX));
            this.LineToLineY = this.query.AddColumn(SrcConstants.PageLayoutLineToLineY, nameof(SrcConstants.PageLayoutLineToLineY));
            this.LineToNodeX = this.query.AddColumn(SrcConstants.PageLayoutLineToNodeX, nameof(SrcConstants.PageLayoutLineToNodeX));
            this.LineToNodeY = this.query.AddColumn(SrcConstants.PageLayoutLineToNodeY, nameof(SrcConstants.PageLayoutLineToNodeY));
            this.LineJumpDirX = this.query.AddColumn(SrcConstants.PageLayoutLineJumpDirX, nameof(SrcConstants.PageLayoutLineJumpDirX));
            this.LineJumpDirY = this.query.AddColumn(SrcConstants.PageLayoutLineJumpDirY, nameof(SrcConstants.PageLayoutLineJumpDirY));
            this.ShapeSplit = this.query.AddColumn(SrcConstants.PageLayoutShapeSplit, nameof(SrcConstants.PageLayoutShapeSplit));
            this.PlaceDepth = this.query.AddColumn(SrcConstants.PageLayoutPlaceDepth, nameof(SrcConstants.PageLayoutPlaceDepth));
            this.PlaceFlip = this.query.AddColumn(SrcConstants.PageLayoutPlaceFlip, nameof(SrcConstants.PageLayoutPlaceFlip));
            this.PlaceStyle = this.query.AddColumn(SrcConstants.PageLayoutPlaceStyle, nameof(SrcConstants.PageLayoutPlaceStyle));
            this.PlowCode = this.query.AddColumn(SrcConstants.PageLayoutPlowCode, nameof(SrcConstants.PageLayoutPlowCode));
            this.ResizePage = this.query.AddColumn(SrcConstants.PageLayoutResizePage, nameof(SrcConstants.PageLayoutResizePage));
            this.RouteStyle = this.query.AddColumn(SrcConstants.PageLayoutRouteStyle, nameof(SrcConstants.PageLayoutRouteStyle));
            this.AvoidPageBreaks = this.query.AddColumn(SrcConstants.PageLayoutAvoidPageBreaks, nameof(SrcConstants.PageLayoutAvoidPageBreaks));
        }


        public override Pages.PageLayoutCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Pages.PageLayoutCells();
            cells.AvenueSizeX = row[this.AvenueSizeX];
            cells.AvenueSizeY = row[this.AvenueSizeY];
            cells.BlockSizeX = row[this.BlockSizeX];
            cells.BlockSizeY = row[this.BlockSizeY];
            cells.CtrlAsInput = row[this.ControlAsInput];
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
            cells.LineJumpDirX = row[this.LineJumpDirX];
            cells.LineJumpDirY = row[this.LineJumpDirY];
            cells.PageShapeSplit = row[this.ShapeSplit];
            cells.PlaceDepth = row[this.PlaceDepth];
            cells.PlaceFlip = row[this.PlaceFlip];
            cells.PlaceStyle = row[this.PlaceStyle];
            cells.PlowCode = row[this.PlowCode];
            cells.ResizePage = row[this.ResizePage];
            cells.RouteStyle = row[this.RouteStyle];
            cells.AvoidPageBreaks = row[this.AvoidPageBreaks];
            return cells;
        }
    }
}