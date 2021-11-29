using VASS=VisioAutomation.ShapeSheet;

namespace VisioAutomation.Models.Dom
{
    public class PageCells
    {
        // PageLayout
        public Core.CellValue PageLayoutAvenueSizeX { get; set; }
        public Core.CellValue PageLayoutAvenueSizeY { get; set; }
        public Core.CellValue PageLayoutBlockSizeX { get; set; }
        public Core.CellValue PageLayoutBlockSizeY { get; set; }
        public Core.CellValue PageLayoutControlAsInput { get; set; }
        public Core.CellValue PageLayoutDynamicsOff { get; set; }
        public Core.CellValue PageLayoutEnableGrid { get; set; }
        public Core.CellValue PageLayoutLineAdjustFrom { get; set; }
        public Core.CellValue PageLayoutLineAdjustTo { get; set; }
        public Core.CellValue PageLayoutLineJumpCode { get; set; }
        public Core.CellValue PageLayoutLineJumpFactorX { get; set; }
        public Core.CellValue PageLayoutLineJumpFactorY { get; set; }
        public Core.CellValue PageLayoutLineJumpStyle { get; set; }
        public Core.CellValue PageLayoutLineRouteExt { get; set; }
        public Core.CellValue PageLayoutLineToLineX { get; set; }
        public Core.CellValue PageLayoutLineToLineY { get; set; }
        public Core.CellValue PageLayoutLineToNodeX { get; set; }
        public Core.CellValue PageLayoutLineToNodeY { get; set; }
        public Core.CellValue PageLayoutPageLineJumpDirX { get; set; }
        public Core.CellValue PageLayoutPageLineJumpDirY { get; set; }
        public Core.CellValue PageShapeSplit { get; set; }
        public Core.CellValue PlaceDepth { get; set; }
        public Core.CellValue PlaceFlip { get; set; }
        public Core.CellValue PlaceStyle { get; set; }
        public Core.CellValue PlowCode { get; set; }
        public Core.CellValue ResizePage { get; set; }
        public Core.CellValue RouteStyle { get; set; }

        public void Apply(VASS.Writers.SidSrcWriter writer, short id)
        {
            writer.SetValue(id, Core.SrcConstants.PageLayoutAvenueSizeX, this.PageLayoutAvenueSizeX);
            writer.SetValue(id, Core.SrcConstants.PageLayoutAvenueSizeY, this.PageLayoutAvenueSizeY);
            writer.SetValue(id, Core.SrcConstants.PageLayoutBlockSizeX, this.PageLayoutBlockSizeX);
            writer.SetValue(id, Core.SrcConstants.PageLayoutBlockSizeY, this.PageLayoutBlockSizeY);
            writer.SetValue(id, Core.SrcConstants.PageLayoutControlAsInput, this.PageLayoutControlAsInput);
            writer.SetValue(id, Core.SrcConstants.PageLayoutDynamicsOff, this.PageLayoutDynamicsOff);
            writer.SetValue(id, Core.SrcConstants.PageLayoutEnableGrid, this.PageLayoutEnableGrid);
            writer.SetValue(id, Core.SrcConstants.PageLayoutLineAdjustFrom, this.PageLayoutLineAdjustFrom);
            writer.SetValue(id, Core.SrcConstants.PageLayoutLineAdjustTo, this.PageLayoutLineAdjustTo);
            writer.SetValue(id, Core.SrcConstants.PageLayoutLineJumpCode, this.PageLayoutLineJumpCode);
            writer.SetValue(id, Core.SrcConstants.PageLayoutLineJumpFactorX, this.PageLayoutLineJumpFactorX);
            writer.SetValue(id, Core.SrcConstants.PageLayoutLineJumpFactorY, this.PageLayoutLineJumpFactorY);
            writer.SetValue(id, Core.SrcConstants.PageLayoutLineJumpStyle, this.PageLayoutLineJumpStyle);
            writer.SetValue(id, Core.SrcConstants.PageLayoutLineRouteExt, this.PageLayoutLineRouteExt);
            writer.SetValue(id, Core.SrcConstants.PageLayoutLineToLineX, this.PageLayoutLineToLineX);
            writer.SetValue(id, Core.SrcConstants.PageLayoutLineToLineY, this.PageLayoutLineToLineY);
            writer.SetValue(id, Core.SrcConstants.PageLayoutLineToNodeX, this.PageLayoutLineToNodeX);
            writer.SetValue(id, Core.SrcConstants.PageLayoutLineToNodeY, this.PageLayoutLineToNodeY);
            writer.SetValue(id, Core.SrcConstants.PageLayoutLineJumpDirX, this.PageLayoutPageLineJumpDirX);
            writer.SetValue(id, Core.SrcConstants.PageLayoutLineJumpDirY, this.PageLayoutPageLineJumpDirY);
            writer.SetValue(id, Core.SrcConstants.PageLayoutShapeSplit, this.PageShapeSplit);
            writer.SetValue(id, Core.SrcConstants.PageLayoutPlaceDepth, this.PlaceDepth);
            writer.SetValue(id, Core.SrcConstants.PageLayoutPlaceFlip, this.PlaceFlip);
            writer.SetValue(id, Core.SrcConstants.PageLayoutPlaceStyle, this.PlaceStyle);
            writer.SetValue(id, Core.SrcConstants.PageLayoutPlowCode, this.PlowCode);
            writer.SetValue(id, Core.SrcConstants.PageLayoutResizePage, this.ResizePage);
            writer.SetValue(id, Core.SrcConstants.PageLayoutRouteStyle, this.RouteStyle);
        }
    }
}