using VASS=VisioAutomation.ShapeSheet;

namespace VisioAutomation.Models.Dom
{
    public class PageCells
    {
        // PageLayout
        public ShapeSheet.CellValue PageLayoutAvenueSizeX { get; set; }
        public ShapeSheet.CellValue PageLayoutAvenueSizeY { get; set; }
        public ShapeSheet.CellValue PageLayoutBlockSizeX { get; set; }
        public ShapeSheet.CellValue PageLayoutBlockSizeY { get; set; }
        public ShapeSheet.CellValue PageLayoutControlAsInput { get; set; }
        public ShapeSheet.CellValue PageLayoutDynamicsOff { get; set; }
        public ShapeSheet.CellValue PageLayoutEnableGrid { get; set; }
        public ShapeSheet.CellValue PageLayoutLineAdjustFrom { get; set; }
        public ShapeSheet.CellValue PageLayoutLineAdjustTo { get; set; }
        public ShapeSheet.CellValue PageLayoutLineJumpCode { get; set; }
        public ShapeSheet.CellValue PageLayoutLineJumpFactorX { get; set; }
        public ShapeSheet.CellValue PageLayoutLineJumpFactorY { get; set; }
        public ShapeSheet.CellValue PageLayoutLineJumpStyle { get; set; }
        public ShapeSheet.CellValue PageLayoutLineRouteExt { get; set; }
        public ShapeSheet.CellValue PageLayoutLineToLineX { get; set; }
        public ShapeSheet.CellValue PageLayoutLineToLineY { get; set; }
        public ShapeSheet.CellValue PageLayoutLineToNodeX { get; set; }
        public ShapeSheet.CellValue PageLayoutLineToNodeY { get; set; }
        public ShapeSheet.CellValue PageLayoutPageLineJumpDirX { get; set; }
        public ShapeSheet.CellValue PageLayoutPageLineJumpDirY { get; set; }
        public ShapeSheet.CellValue PageShapeSplit { get; set; }
        public ShapeSheet.CellValue PlaceDepth { get; set; }
        public ShapeSheet.CellValue PlaceFlip { get; set; }
        public ShapeSheet.CellValue PlaceStyle { get; set; }
        public ShapeSheet.CellValue PlowCode { get; set; }
        public ShapeSheet.CellValue ResizePage { get; set; }
        public ShapeSheet.CellValue RouteStyle { get; set; }

        public void Apply(VASS.Writers.SidSrcWriter writer, short id)
        {
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutAvenueSizeX, this.PageLayoutAvenueSizeX);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutAvenueSizeY, this.PageLayoutAvenueSizeY);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutBlockSizeX, this.PageLayoutBlockSizeX);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutBlockSizeY, this.PageLayoutBlockSizeY);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutControlAsInput, this.PageLayoutControlAsInput);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutDynamicsOff, this.PageLayoutDynamicsOff);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutEnableGrid, this.PageLayoutEnableGrid);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutLineAdjustFrom, this.PageLayoutLineAdjustFrom);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutLineAdjustTo, this.PageLayoutLineAdjustTo);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutLineJumpCode, this.PageLayoutLineJumpCode);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutLineJumpFactorX, this.PageLayoutLineJumpFactorX);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutLineJumpFactorY, this.PageLayoutLineJumpFactorY);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutLineJumpStyle, this.PageLayoutLineJumpStyle);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutLineRouteExt, this.PageLayoutLineRouteExt);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutLineToLineX, this.PageLayoutLineToLineX);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutLineToLineY, this.PageLayoutLineToLineY);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutLineToNodeX, this.PageLayoutLineToNodeX);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutLineToNodeY, this.PageLayoutLineToNodeY);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutLineJumpDirX, this.PageLayoutPageLineJumpDirX);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutLineJumpDirY, this.PageLayoutPageLineJumpDirY);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutShapeSplit, this.PageShapeSplit);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutPlaceDepth, this.PlaceDepth);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutPlaceFlip, this.PlaceFlip);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutPlaceStyle, this.PlaceStyle);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutPlowCode, this.PlowCode);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutResizePage, this.ResizePage);
            writer.SetValue(id, ShapeSheet.SrcConstants.PageLayoutRouteStyle, this.RouteStyle);
        }
    }
}