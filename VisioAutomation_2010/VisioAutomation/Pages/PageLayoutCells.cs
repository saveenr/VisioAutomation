using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;

namespace VisioAutomation.Pages
{
    public class PageLayoutCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public ShapeSheet.CellData AvenueSizeX { get; set; }
        public ShapeSheet.CellData AvenueSizeY { get; set; }
        public ShapeSheet.CellData BlockSizeX { get; set; }
        public ShapeSheet.CellData BlockSizeY { get; set; }
        public ShapeSheet.CellData CtrlAsInput { get; set; }
        public ShapeSheet.CellData DynamicsOff { get; set; }
        public ShapeSheet.CellData EnableGrid { get; set; }
        public ShapeSheet.CellData LineAdjustFrom { get; set; }
        public ShapeSheet.CellData LineAdjustTo { get; set; }
        public ShapeSheet.CellData LineJumpCode { get; set; }
        public ShapeSheet.CellData LineJumpFactorX { get; set; }
        public ShapeSheet.CellData LineJumpFactorY { get; set; }
        public ShapeSheet.CellData LineJumpStyle { get; set; }
        public ShapeSheet.CellData LineRouteExt { get; set; }
        public ShapeSheet.CellData LineToLineX { get; set; }
        public ShapeSheet.CellData LineToLineY { get; set; }
        public ShapeSheet.CellData LineToNodeX { get; set; }
        public ShapeSheet.CellData LineToNodeY { get; set; }
        public ShapeSheet.CellData LineJumpDirX { get; set; }
        public ShapeSheet.CellData LineJumpDirY { get; set; }
        public ShapeSheet.CellData PageShapeSplit { get; set; }
        public ShapeSheet.CellData PlaceDepth { get; set; }
        public ShapeSheet.CellData PlaceFlip { get; set; }
        public ShapeSheet.CellData PlaceStyle { get; set; }
        public ShapeSheet.CellData PlowCode { get; set; }
        public ShapeSheet.CellData ResizePage { get; set; }
        public ShapeSheet.CellData RouteStyle { get; set; }
        public ShapeSheet.CellData AvoidPageBreaks { get; set; } // new in visio 2010

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutAvenueSizeX, this.AvenueSizeX.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutAvenueSizeY, this.AvenueSizeY.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutBlockSizeX, this.BlockSizeX.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutBlockSizeY, this.BlockSizeY.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutControlAsInput, this.CtrlAsInput.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutDynamicsOff, this.DynamicsOff.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutEnableGrid, this.EnableGrid.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineAdjustFrom, this.LineAdjustFrom.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineAdjustTo, this.LineAdjustTo.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpCode, this.LineJumpCode.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpFactorX, this.LineJumpFactorX.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpFactorY, this.LineJumpFactorY.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpStyle, this.LineJumpStyle.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineRouteExt, this.LineRouteExt.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineToLineX, this.LineToLineX.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineToLineY, this.LineToLineY.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineToNodeX, this.LineToNodeX.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineToNodeY, this.LineToNodeY.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpDirX, this.LineJumpDirX.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpDirY, this.LineJumpDirY.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutShapeSplit, this.PageShapeSplit.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlaceDepth, this.PlaceDepth.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlaceFlip, this.PlaceFlip.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlaceStyle, this.PlaceStyle.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlowCode, this.PlowCode.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutResizePage, this.ResizePage.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutRouteStyle, this.RouteStyle.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutAvoidPageBreaks, this.AvoidPageBreaks.Value);
            }
        }

        public static PageLayoutCells GetCells(Microsoft.Office.Interop.Visio.Shape shape, ShapeSheet.CellValueType cvt)
        {
            var query = PageLayoutCells.lazy_query.Value;
            return query.GetCellGroup(shape,cvt);
        }

        private static readonly System.Lazy<PageLayoutCellsReader> lazy_query = new System.Lazy<PageLayoutCellsReader>();
    }
}