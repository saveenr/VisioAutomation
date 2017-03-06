using VisioAutomation.ShapeSheet;
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
        public ShapeSheet.CellValueLiteral PageLayoutCtrlAsInput { get; set; }
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
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutAvenueSizeX, this.PageLayoutAvenueSizeX);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutAvenueSizeY, this.PageLayoutAvenueSizeY);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutBlockSizeX, this.PageLayoutBlockSizeX);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutBlockSizeY, this.PageLayoutBlockSizeY);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutCtrlAsInput, this.PageLayoutCtrlAsInput);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutDynamicsOff, this.PageLayoutDynamicsOff);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutEnableGrid, this.PageLayoutEnableGrid);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutLineAdjustFrom, this.PageLayoutLineAdjustFrom);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutLineAdjustTo, this.PageLayoutLineAdjustTo);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutLineJumpCode, this.PageLayoutLineJumpCode);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutLineJumpFactorX, this.PageLayoutLineJumpFactorX);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutLineJumpFactorY, this.PageLayoutLineJumpFactorY);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutLineJumpStyle, this.PageLayoutLineJumpStyle);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutLineRouteExt, this.PageLayoutLineRouteExt);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutLineToLineX, this.PageLayoutLineToLineX);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutLineToLineY, this.PageLayoutLineToLineY);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutLineToNodeX, this.PageLayoutLineToNodeX);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutLineToNodeY, this.PageLayoutLineToNodeY);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutLineJumpDirX, this.PageLayoutPageLineJumpDirX);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutLineJumpDirY, this.PageLayoutPageLineJumpDirY);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutPageShapeSplit, this.PageShapeSplit);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutPlaceDepth, this.PlaceDepth);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutPlaceFlip, this.PlaceFlip);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutPlaceStyle, this.PlaceStyle);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutPlowCode, this.PlowCode);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutResizePage, this.ResizePage);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLayoutRouteStyle, this.RouteStyle);
        }
    }
}