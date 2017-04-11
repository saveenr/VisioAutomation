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
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutAvenueSizeX, this.AvenueSizeX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutAvenueSizeY, this.AvenueSizeY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutBlockSizeX, this.BlockSizeX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutBlockSizeY, this.BlockSizeY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutControlAsInput, this.CtrlAsInput.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutDynamicsOff, this.DynamicsOff.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutEnableGrid, this.EnableGrid.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineAdjustFrom, this.LineAdjustFrom.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineAdjustTo, this.LineAdjustTo.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpCode, this.LineJumpCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpFactorX, this.LineJumpFactorX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpFactorY, this.LineJumpFactorY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpStyle, this.LineJumpStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineRouteExt, this.LineRouteExt.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineToLineX, this.LineToLineX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineToLineY, this.LineToLineY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineToNodeX, this.LineToNodeX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineToNodeY, this.LineToNodeY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpDirX, this.LineJumpDirX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutLineJumpDirY, this.LineJumpDirY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutShapeSplit, this.PageShapeSplit.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlaceDepth, this.PlaceDepth.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlaceFlip, this.PlaceFlip.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlaceStyle, this.PlaceStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutPlowCode, this.PlowCode.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutResizePage, this.ResizePage.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutRouteStyle, this.RouteStyle.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.PageLayoutAvoidPageBreaks, this.AvoidPageBreaks.Formula);
            }
        }

        public static PageLayoutCells GetCells(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = PageLayoutCells.lazy_query.Value;
            return query.GetCellGroup(shape);
        }

        private static readonly System.Lazy<PageLayoutCellsReader> lazy_query = new System.Lazy<PageLayoutCellsReader>();
    }
}