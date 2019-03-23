using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.Models.Dom
{
    public class PageCells
    {
        // PageLayout
        public ShapeSheet.CellValueLiteral PageLayoutAvenueSizeX { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutAvenueSizeY { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutBlockSizeX { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutBlockSizeY { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutControlAsInput { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutDynamicsOff { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutEnableGrid { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutLineAdjustFrom { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutLineAdjustTo { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutLineJumpCode { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutLineJumpFactorX { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutLineJumpFactorY { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutLineJumpStyle { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutLineRouteExt { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutLineToLineX { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutLineToLineY { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutLineToNodeX { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutLineToNodeY { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutPageLineJumpDirX { get; set; }
        public ShapeSheet.CellValueLiteral PageLayoutPageLineJumpDirY { get; set; }
        public ShapeSheet.CellValueLiteral PageShapeSplit { get; set; }
        public ShapeSheet.CellValueLiteral PlaceDepth { get; set; }
        public ShapeSheet.CellValueLiteral PlaceFlip { get; set; }
        public ShapeSheet.CellValueLiteral PlaceStyle { get; set; }
        public ShapeSheet.CellValueLiteral PlowCode { get; set; }
        public ShapeSheet.CellValueLiteral ResizePage { get; set; }
        public ShapeSheet.CellValueLiteral RouteStyle { get; set; }

        public void Apply(SidSrcWriter writer, short id)
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