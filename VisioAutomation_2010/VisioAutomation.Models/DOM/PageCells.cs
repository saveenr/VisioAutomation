using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.Models.Dom
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

        public void Apply(FormulaWriterSIDSRC writer, short id)
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