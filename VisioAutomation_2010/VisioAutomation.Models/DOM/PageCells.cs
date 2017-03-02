using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Models.Dom
{
    public class PageCells
    {
        // PageLayout
        public ShapeSheet.CellValueLiteral AvenueSizeX { get; set; }
        public ShapeSheet.CellValueLiteral AvenueSizeY { get; set; }
        public ShapeSheet.CellValueLiteral BlockSizeX { get; set; }
        public ShapeSheet.CellValueLiteral BlockSizeY { get; set; }
        public ShapeSheet.CellValueLiteral CtrlAsInput { get; set; }
        public ShapeSheet.CellValueLiteral DynamicsOff { get; set; }
        public ShapeSheet.CellValueLiteral EnableGrid { get; set; }
        public ShapeSheet.CellValueLiteral LineAdjustFrom { get; set; }
        public ShapeSheet.CellValueLiteral LineAdjustTo { get; set; }
        public ShapeSheet.CellValueLiteral LineJumpCode { get; set; }
        public ShapeSheet.CellValueLiteral LineJumpFactorX { get; set; }
        public ShapeSheet.CellValueLiteral LineJumpFactorY { get; set; }
        public ShapeSheet.CellValueLiteral LineJumpStyle { get; set; }
        public ShapeSheet.CellValueLiteral LineRouteExt { get; set; }
        public ShapeSheet.CellValueLiteral LineToLineX { get; set; }
        public ShapeSheet.CellValueLiteral LineToLineY { get; set; }
        public ShapeSheet.CellValueLiteral LineToNodeX { get; set; }
        public ShapeSheet.CellValueLiteral LineToNodeY { get; set; }
        public ShapeSheet.CellValueLiteral PageLineJumpDirX { get; set; }
        public ShapeSheet.CellValueLiteral PageLineJumpDirY { get; set; }
        public ShapeSheet.CellValueLiteral PageShapeSplit { get; set; }
        public ShapeSheet.CellValueLiteral PlaceDepth { get; set; }
        public ShapeSheet.CellValueLiteral PlaceFlip { get; set; }
        public ShapeSheet.CellValueLiteral PlaceStyle { get; set; }
        public ShapeSheet.CellValueLiteral PlowCode { get; set; }
        public ShapeSheet.CellValueLiteral ResizePage { get; set; }
        public ShapeSheet.CellValueLiteral RouteStyle { get; set; }

        public void Apply(ShapeSheetWriter writer, short id)
        {
            writer.SetFormula(id, ShapeSheet.SrcConstants.AvenueSizeX, this.AvenueSizeX);
            writer.SetFormula(id, ShapeSheet.SrcConstants.AvenueSizeY, this.AvenueSizeY);
            writer.SetFormula(id, ShapeSheet.SrcConstants.BlockSizeX, this.BlockSizeX);
            writer.SetFormula(id, ShapeSheet.SrcConstants.BlockSizeY, this.BlockSizeY);
            writer.SetFormula(id, ShapeSheet.SrcConstants.CtrlAsInput, this.CtrlAsInput);
            writer.SetFormula(id, ShapeSheet.SrcConstants.DynamicsOff, this.DynamicsOff);
            writer.SetFormula(id, ShapeSheet.SrcConstants.EnableGrid, this.EnableGrid);
            writer.SetFormula(id, ShapeSheet.SrcConstants.LineAdjustFrom, this.LineAdjustFrom);
            writer.SetFormula(id, ShapeSheet.SrcConstants.LineAdjustTo, this.LineAdjustTo);
            writer.SetFormula(id, ShapeSheet.SrcConstants.LineJumpCode, this.LineJumpCode);
            writer.SetFormula(id, ShapeSheet.SrcConstants.LineJumpFactorX, this.LineJumpFactorX);
            writer.SetFormula(id, ShapeSheet.SrcConstants.LineJumpFactorY, this.LineJumpFactorY);
            writer.SetFormula(id, ShapeSheet.SrcConstants.LineJumpStyle, this.LineJumpStyle);
            writer.SetFormula(id, ShapeSheet.SrcConstants.LineRouteExt, this.LineRouteExt);
            writer.SetFormula(id, ShapeSheet.SrcConstants.LineToLineX, this.LineToLineX);
            writer.SetFormula(id, ShapeSheet.SrcConstants.LineToLineY, this.LineToLineY);
            writer.SetFormula(id, ShapeSheet.SrcConstants.LineToNodeX, this.LineToNodeX);
            writer.SetFormula(id, ShapeSheet.SrcConstants.LineToNodeY, this.LineToNodeY);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLineJumpDirX, this.PageLineJumpDirX);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageLineJumpDirY, this.PageLineJumpDirY);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PageShapeSplit, this.PageShapeSplit);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PlaceDepth, this.PlaceDepth);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PlaceFlip, this.PlaceFlip);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PlaceStyle, this.PlaceStyle);
            writer.SetFormula(id, ShapeSheet.SrcConstants.PlowCode, this.PlowCode);
            writer.SetFormula(id, ShapeSheet.SrcConstants.ResizePage, this.ResizePage);
            writer.SetFormula(id, ShapeSheet.SrcConstants.RouteStyle, this.RouteStyle);
        }
    }
}