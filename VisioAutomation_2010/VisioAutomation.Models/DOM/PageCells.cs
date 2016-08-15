using VisioAutomation.ShapeSheet.Writers;

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

        public void Apply(FormulaWriterSIDSRC update, short id)
        {
            update.SetFormula(id, ShapeSheet.SRCConstants.AvenueSizeX, this.AvenueSizeX);
            update.SetFormula(id, ShapeSheet.SRCConstants.AvenueSizeY, this.AvenueSizeY);
            update.SetFormula(id, ShapeSheet.SRCConstants.BlockSizeX, this.BlockSizeX);
            update.SetFormula(id, ShapeSheet.SRCConstants.BlockSizeY, this.BlockSizeY);
            update.SetFormula(id, ShapeSheet.SRCConstants.CtrlAsInput, this.CtrlAsInput);
            update.SetFormula(id, ShapeSheet.SRCConstants.DynamicsOff, this.DynamicsOff);
            update.SetFormula(id, ShapeSheet.SRCConstants.EnableGrid, this.EnableGrid);
            update.SetFormula(id, ShapeSheet.SRCConstants.LineAdjustFrom, this.LineAdjustFrom);
            update.SetFormula(id, ShapeSheet.SRCConstants.LineAdjustTo, this.LineAdjustTo);
            update.SetFormula(id, ShapeSheet.SRCConstants.LineJumpCode, this.LineJumpCode);
            update.SetFormula(id, ShapeSheet.SRCConstants.LineJumpFactorX, this.LineJumpFactorX);
            update.SetFormula(id, ShapeSheet.SRCConstants.LineJumpFactorY, this.LineJumpFactorY);
            update.SetFormula(id, ShapeSheet.SRCConstants.LineJumpStyle, this.LineJumpStyle);
            update.SetFormula(id, ShapeSheet.SRCConstants.LineRouteExt, this.LineRouteExt);
            update.SetFormula(id, ShapeSheet.SRCConstants.LineToLineX, this.LineToLineX);
            update.SetFormula(id, ShapeSheet.SRCConstants.LineToLineY, this.LineToLineY);
            update.SetFormula(id, ShapeSheet.SRCConstants.LineToNodeX, this.LineToNodeX);
            update.SetFormula(id, ShapeSheet.SRCConstants.LineToNodeY, this.LineToNodeY);
            update.SetFormula(id, ShapeSheet.SRCConstants.PageLineJumpDirX, this.PageLineJumpDirX);
            update.SetFormula(id, ShapeSheet.SRCConstants.PageLineJumpDirY, this.PageLineJumpDirY);
            update.SetFormula(id, ShapeSheet.SRCConstants.PageShapeSplit, this.PageShapeSplit);
            update.SetFormula(id, ShapeSheet.SRCConstants.PlaceDepth, this.PlaceDepth);
            update.SetFormula(id, ShapeSheet.SRCConstants.PlaceFlip, this.PlaceFlip);
            update.SetFormula(id, ShapeSheet.SRCConstants.PlaceStyle, this.PlaceStyle);
            update.SetFormula(id, ShapeSheet.SRCConstants.PlowCode, this.PlowCode);
            update.SetFormula(id, ShapeSheet.SRCConstants.ResizePage, this.ResizePage);
            update.SetFormula(id, ShapeSheet.SRCConstants.RouteStyle, this.RouteStyle);
        }
    }
}