using VisioAutomation.ShapeSheet.Update;

namespace VisioAutomation.DOM
{
    public class PageCells
    {
        // PageLayout
        public ShapeSheet.FormulaLiteral AvenueSizeX { get; set; }
        public ShapeSheet.FormulaLiteral AvenueSizeY { get; set; }
        public ShapeSheet.FormulaLiteral BlockSizeX { get; set; }
        public ShapeSheet.FormulaLiteral BlockSizeY { get; set; }
        public ShapeSheet.FormulaLiteral CtrlAsInput { get; set; }
        public ShapeSheet.FormulaLiteral DynamicsOff { get; set; }
        public ShapeSheet.FormulaLiteral EnableGrid { get; set; }
        public ShapeSheet.FormulaLiteral LineAdjustFrom { get; set; }
        public ShapeSheet.FormulaLiteral LineAdjustTo { get; set; }
        public ShapeSheet.FormulaLiteral LineJumpCode { get; set; }
        public ShapeSheet.FormulaLiteral LineJumpFactorX { get; set; }
        public ShapeSheet.FormulaLiteral LineJumpFactorY { get; set; }
        public ShapeSheet.FormulaLiteral LineJumpStyle { get; set; }
        public ShapeSheet.FormulaLiteral LineRouteExt { get; set; }
        public ShapeSheet.FormulaLiteral LineToLineX { get; set; }
        public ShapeSheet.FormulaLiteral LineToLineY { get; set; }
        public ShapeSheet.FormulaLiteral LineToNodeX { get; set; }
        public ShapeSheet.FormulaLiteral LineToNodeY { get; set; }
        public ShapeSheet.FormulaLiteral PageLineJumpDirX { get; set; }
        public ShapeSheet.FormulaLiteral PageLineJumpDirY { get; set; }
        public ShapeSheet.FormulaLiteral PageShapeSplit { get; set; }
        public ShapeSheet.FormulaLiteral PlaceDepth { get; set; }
        public ShapeSheet.FormulaLiteral PlaceFlip { get; set; }
        public ShapeSheet.FormulaLiteral PlaceStyle { get; set; }
        public ShapeSheet.FormulaLiteral PlowCode { get; set; }
        public ShapeSheet.FormulaLiteral ResizePage { get; set; }
        public ShapeSheet.FormulaLiteral RouteStyle { get; set; }

        public void Apply(UpdateSIDSRC update, short id)
        {
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.AvenueSizeX, this.AvenueSizeX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.AvenueSizeY, this.AvenueSizeY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.BlockSizeX, this.BlockSizeX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.BlockSizeY, this.BlockSizeY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.CtrlAsInput, this.CtrlAsInput);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.DynamicsOff, this.DynamicsOff);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.EnableGrid, this.EnableGrid);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineAdjustFrom, this.LineAdjustFrom);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineAdjustTo, this.LineAdjustTo);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineJumpCode, this.LineJumpCode);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineJumpFactorX, this.LineJumpFactorX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineJumpFactorY, this.LineJumpFactorY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineJumpStyle, this.LineJumpStyle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineRouteExt, this.LineRouteExt);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineToLineX, this.LineToLineX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineToLineY, this.LineToLineY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineToNodeX, this.LineToNodeX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineToNodeY, this.LineToNodeY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PageLineJumpDirX, this.PageLineJumpDirX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PageLineJumpDirY, this.PageLineJumpDirY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PageShapeSplit, this.PageShapeSplit);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PlaceDepth, this.PlaceDepth);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PlaceFlip, this.PlaceFlip);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PlaceStyle, this.PlaceStyle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PlowCode, this.PlowCode);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ResizePage, this.ResizePage);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.RouteStyle, this.RouteStyle);
        }
    }
}