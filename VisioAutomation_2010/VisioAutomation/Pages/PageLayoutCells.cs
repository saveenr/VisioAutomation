using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;

namespace VisioAutomation.Pages
{
    public class PageLayoutCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral AvenueSizeX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral AvenueSizeY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral BlockSizeX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral BlockSizeY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral CtrlAsInput { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral DynamicsOff { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral EnableGrid { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineAdjustFrom { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineAdjustTo { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpCode { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpFactorX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpFactorY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpStyle { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineRouteExt { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineToLineX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineToLineY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineToNodeX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineToNodeY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpDirX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpDirY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PageShapeSplit { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PlaceDepth { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PlaceFlip { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PlaceStyle { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PlowCode { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ResizePage { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral RouteStyle { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral AvoidPageBreaks { get; set; } // new in visio 2010

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutAvenueSizeX, this.AvenueSizeX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutAvenueSizeY, this.AvenueSizeY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutBlockSizeX, this.BlockSizeX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutBlockSizeY, this.BlockSizeY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutControlAsInput, this.CtrlAsInput.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutDynamicsOff, this.DynamicsOff.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutEnableGrid, this.EnableGrid.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineAdjustFrom, this.LineAdjustFrom.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineAdjustTo, this.LineAdjustTo.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineJumpCode, this.LineJumpCode.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineJumpFactorX, this.LineJumpFactorX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineJumpFactorY, this.LineJumpFactorY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineJumpStyle, this.LineJumpStyle.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineRouteExt, this.LineRouteExt.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineToLineX, this.LineToLineX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineToLineY, this.LineToLineY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineToNodeX, this.LineToNodeX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineToNodeY, this.LineToNodeY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineJumpDirX, this.LineJumpDirX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutLineJumpDirY, this.LineJumpDirY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutShapeSplit, this.PageShapeSplit.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutPlaceDepth, this.PlaceDepth.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutPlaceFlip, this.PlaceFlip.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutPlaceStyle, this.PlaceStyle.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutPlowCode, this.PlowCode.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutResizePage, this.ResizePage.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutRouteStyle, this.RouteStyle.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.PageLayoutAvoidPageBreaks, this.AvoidPageBreaks.Value);
            }
        }

        public static PageLayoutCells GetFormulas(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = PageLayoutCells.lazy_query.Value;
            return query.GetFormulas(shape);
        }

        public static PageLayoutCells GetResults(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = PageLayoutCells.lazy_query.Value;
            return query.GetResults(shape);
        }
        private static readonly System.Lazy<PageLayoutCellsReader> lazy_query = new System.Lazy<PageLayoutCellsReader>();
    }
}