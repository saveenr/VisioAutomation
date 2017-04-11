using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Pages
{
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
            cells.AvenueSizeX = row[this.PageLayoutAvenueSizeX];
            cells.AvenueSizeY = row[this.PageLayoutAvenueSizeY];
            cells.BlockSizeX = row[this.PageLayoutBlockSizeX];
            cells.BlockSizeY = row[this.PageLayoutBlockSizeY];
            cells.CtrlAsInput = row[this.PageLayoutControlAsInput];
            cells.DynamicsOff = row[this.PageLayoutDynamicsOff];
            cells.EnableGrid = row[this.PageLayoutEnableGrid];
            cells.LineAdjustFrom = row[this.PageLayoutLineAdjustFrom];
            cells.LineAdjustTo = row[this.PageLayoutLineAdjustTo];
            cells.LineJumpCode = row[this.PageLayoutLineJumpCode];
            cells.LineJumpFactorX = row[this.PageLayoutLineJumpFactorX];
            cells.LineJumpFactorY = row[this.PageLayoutLineJumpFactorY];
            cells.LineJumpStyle = row[this.PageLayoutLineJumpStyle];
            cells.LineRouteExt = row[this.PageLayoutLineRouteExt];
            cells.LineToLineX = row[this.PageLayoutLineToLineX];
            cells.LineToLineY = row[this.PageLayoutLineToLineY];
            cells.LineToNodeX = row[this.PageLayoutLineToNodeX];
            cells.LineToNodeY = row[this.PageLayoutLineToNodeY];
            cells.LineJumpDirX = row[this.PageLayoutLineJumpDirX];
            cells.LineJumpDirY = row[this.PageLayoutLineJumpDirY];
            cells.PageShapeSplit = row[this.PageLayoutShapeSplit];
            cells.PlaceDepth = row[this.PageLayoutPlaceDepth];
            cells.PlaceFlip = row[this.PageLayoutPlaceFlip];
            cells.PlaceStyle = row[this.PageLayoutPlaceStyle];
            cells.PlowCode = row[this.PageLayoutPlowCode];
            cells.ResizePage = row[this.PageLayoutResizePage];
            cells.RouteStyle = row[this.PageLayoutRouteStyle];
            cells.AvoidPageBreaks = row[this.PageLayoutAvoidPageBreaks];
            return cells;
        }
    }
}