using System.Collections.Generic;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

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

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short id)
        {
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.AvenueSizeX, AvenueSizeX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.AvenueSizeY, AvenueSizeY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.BlockSizeX, BlockSizeX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.BlockSizeY, BlockSizeY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.CtrlAsInput, CtrlAsInput);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.DynamicsOff, DynamicsOff);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.EnableGrid, EnableGrid);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineAdjustFrom, LineAdjustFrom);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineAdjustTo, LineAdjustTo);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineJumpCode, LineJumpCode);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineJumpFactorX, LineJumpFactorX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineJumpFactorY, LineJumpFactorY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineJumpStyle, LineJumpStyle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineRouteExt, LineRouteExt);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineToLineX, LineToLineX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineToLineY, LineToLineY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineToNodeX, LineToNodeX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.LineToNodeY, LineToNodeY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PageLineJumpDirX, PageLineJumpDirX);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PageLineJumpDirY, PageLineJumpDirY);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PageShapeSplit, PageShapeSplit);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PlaceDepth, PlaceDepth);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PlaceFlip, PlaceFlip);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PlaceStyle, PlaceStyle);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.PlowCode, PlowCode);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.ResizePage, ResizePage);
            update.SetFormulaIgnoreNull(id, ShapeSheet.SRCConstants.RouteStyle, RouteStyle);
        }

        public PageCells ShallowCopy()
        {
            return (PageCells) this.MemberwiseClone();
        }
    }
}