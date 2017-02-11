using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Models.Dom
{
    public class PageCells
    {
        // PageLayout
        public ShapeSheet.ValueLiteral AvenueSizeX { get; set; }
        public ShapeSheet.ValueLiteral AvenueSizeY { get; set; }
        public ShapeSheet.ValueLiteral BlockSizeX { get; set; }
        public ShapeSheet.ValueLiteral BlockSizeY { get; set; }
        public ShapeSheet.ValueLiteral CtrlAsInput { get; set; }
        public ShapeSheet.ValueLiteral DynamicsOff { get; set; }
        public ShapeSheet.ValueLiteral EnableGrid { get; set; }
        public ShapeSheet.ValueLiteral LineAdjustFrom { get; set; }
        public ShapeSheet.ValueLiteral LineAdjustTo { get; set; }
        public ShapeSheet.ValueLiteral LineJumpCode { get; set; }
        public ShapeSheet.ValueLiteral LineJumpFactorX { get; set; }
        public ShapeSheet.ValueLiteral LineJumpFactorY { get; set; }
        public ShapeSheet.ValueLiteral LineJumpStyle { get; set; }
        public ShapeSheet.ValueLiteral LineRouteExt { get; set; }
        public ShapeSheet.ValueLiteral LineToLineX { get; set; }
        public ShapeSheet.ValueLiteral LineToLineY { get; set; }
        public ShapeSheet.ValueLiteral LineToNodeX { get; set; }
        public ShapeSheet.ValueLiteral LineToNodeY { get; set; }
        public ShapeSheet.ValueLiteral PageLineJumpDirX { get; set; }
        public ShapeSheet.ValueLiteral PageLineJumpDirY { get; set; }
        public ShapeSheet.ValueLiteral PageShapeSplit { get; set; }
        public ShapeSheet.ValueLiteral PlaceDepth { get; set; }
        public ShapeSheet.ValueLiteral PlaceFlip { get; set; }
        public ShapeSheet.ValueLiteral PlaceStyle { get; set; }
        public ShapeSheet.ValueLiteral PlowCode { get; set; }
        public ShapeSheet.ValueLiteral ResizePage { get; set; }
        public ShapeSheet.ValueLiteral RouteStyle { get; set; }

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