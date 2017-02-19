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
            writer.SetFormula(id, ShapeSheet.SRCConstants.AvenueSizeX, this.AvenueSizeX);
            writer.SetFormula(id, ShapeSheet.SRCConstants.AvenueSizeY, this.AvenueSizeY);
            writer.SetFormula(id, ShapeSheet.SRCConstants.BlockSizeX, this.BlockSizeX);
            writer.SetFormula(id, ShapeSheet.SRCConstants.BlockSizeY, this.BlockSizeY);
            writer.SetFormula(id, ShapeSheet.SRCConstants.CtrlAsInput, this.CtrlAsInput);
            writer.SetFormula(id, ShapeSheet.SRCConstants.DynamicsOff, this.DynamicsOff);
            writer.SetFormula(id, ShapeSheet.SRCConstants.EnableGrid, this.EnableGrid);
            writer.SetFormula(id, ShapeSheet.SRCConstants.LineAdjustFrom, this.LineAdjustFrom);
            writer.SetFormula(id, ShapeSheet.SRCConstants.LineAdjustTo, this.LineAdjustTo);
            writer.SetFormula(id, ShapeSheet.SRCConstants.LineJumpCode, this.LineJumpCode);
            writer.SetFormula(id, ShapeSheet.SRCConstants.LineJumpFactorX, this.LineJumpFactorX);
            writer.SetFormula(id, ShapeSheet.SRCConstants.LineJumpFactorY, this.LineJumpFactorY);
            writer.SetFormula(id, ShapeSheet.SRCConstants.LineJumpStyle, this.LineJumpStyle);
            writer.SetFormula(id, ShapeSheet.SRCConstants.LineRouteExt, this.LineRouteExt);
            writer.SetFormula(id, ShapeSheet.SRCConstants.LineToLineX, this.LineToLineX);
            writer.SetFormula(id, ShapeSheet.SRCConstants.LineToLineY, this.LineToLineY);
            writer.SetFormula(id, ShapeSheet.SRCConstants.LineToNodeX, this.LineToNodeX);
            writer.SetFormula(id, ShapeSheet.SRCConstants.LineToNodeY, this.LineToNodeY);
            writer.SetFormula(id, ShapeSheet.SRCConstants.PageLineJumpDirX, this.PageLineJumpDirX);
            writer.SetFormula(id, ShapeSheet.SRCConstants.PageLineJumpDirY, this.PageLineJumpDirY);
            writer.SetFormula(id, ShapeSheet.SRCConstants.PageShapeSplit, this.PageShapeSplit);
            writer.SetFormula(id, ShapeSheet.SRCConstants.PlaceDepth, this.PlaceDepth);
            writer.SetFormula(id, ShapeSheet.SRCConstants.PlaceFlip, this.PlaceFlip);
            writer.SetFormula(id, ShapeSheet.SRCConstants.PlaceStyle, this.PlaceStyle);
            writer.SetFormula(id, ShapeSheet.SRCConstants.PlowCode, this.PlowCode);
            writer.SetFormula(id, ShapeSheet.SRCConstants.ResizePage, this.ResizePage);
            writer.SetFormula(id, ShapeSheet.SRCConstants.RouteStyle, this.RouteStyle);
        }
    }
}