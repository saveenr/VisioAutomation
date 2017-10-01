using System.Collections.Generic;
using VisioAutomation.ShapeSheet;

namespace VisioPowerShell.Models
{
    public class PageLayoutCells : VisioPowerShell.Models.BaseCells
    {
        public string AvenueSizeX;
        public string AvenueSizeY;
        public string AvoidPageBreaks;
        public string BlockSizeX;
        public string BlockSizeY;
        public string ControlAsInput;
        public string DynamicsOff;
        public string EnableGrid;
        public string LineAdjustFrom;
        public string LineAdjustTo;
        public string LineJumpCode;
        public string LineJumpDirX;
        public string LineJumpDirY;
        public string LineJumpFactorX;
        public string LineJumpFactorY;
        public string LineJumpStyle;
        public string LineRouteExt;
        public string LineToLineX;
        public string LineToLineY;
        public string LineToNodeX;
        public string LineToNodeY;
        public string PlaceDepth;
        public string PlaceFlip;
        public string PlaceStyle;
        public string PlowCode;
        public string ResizePage;
        public string RouteStyle;
        public string ShapeSplit;

        public override IEnumerable<CellTuple> GetCellTuples()
        {
            yield return new CellTuple(nameof(SrcConstants.PageLayoutAvenueSizeX), SrcConstants.PageLayoutAvenueSizeX, this.AvenueSizeX);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutAvenueSizeY), SrcConstants.PageLayoutAvenueSizeY, this.AvenueSizeY);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutAvoidPageBreaks), SrcConstants.PageLayoutAvoidPageBreaks, this.AvoidPageBreaks);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutBlockSizeX), SrcConstants.PageLayoutBlockSizeX, this.BlockSizeX);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutBlockSizeY), SrcConstants.PageLayoutBlockSizeY, this.BlockSizeY);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutControlAsInput), SrcConstants.PageLayoutControlAsInput, this.ControlAsInput);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutDynamicsOff), SrcConstants.PageLayoutDynamicsOff, this.DynamicsOff);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutEnableGrid), SrcConstants.PageLayoutEnableGrid, this.EnableGrid);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineAdjustFrom), SrcConstants.PageLayoutLineAdjustFrom, this.LineAdjustFrom);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineAdjustTo), SrcConstants.PageLayoutLineAdjustTo, this.LineAdjustTo);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineJumpCode), SrcConstants.PageLayoutLineJumpCode, this.LineJumpCode);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineJumpDirX), SrcConstants.PageLayoutLineJumpDirX, this.LineJumpDirX);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineJumpDirY), SrcConstants.PageLayoutLineJumpDirY, this.LineJumpDirY);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineJumpFactorX), SrcConstants.PageLayoutLineJumpFactorX, this.LineJumpFactorX);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineJumpFactorY), SrcConstants.PageLayoutLineJumpFactorY, this.LineJumpFactorY);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineJumpStyle), SrcConstants.PageLayoutLineJumpStyle, this.LineJumpStyle);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineRouteExt), SrcConstants.PageLayoutLineRouteExt, this.LineRouteExt);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineToLineX), SrcConstants.PageLayoutLineToLineX, this.LineToLineX);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineToLineY), SrcConstants.PageLayoutLineToLineY, this.LineToLineY);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineToNodeX), SrcConstants.PageLayoutLineToNodeX, this.LineToNodeX);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineToNodeY), SrcConstants.PageLayoutLineToNodeY, this.LineToNodeY);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutPlaceDepth), SrcConstants.PageLayoutPlaceDepth, this.PlaceDepth);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutPlaceFlip), SrcConstants.PageLayoutPlaceFlip, this.PlaceFlip);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutPlaceStyle), SrcConstants.PageLayoutPlaceStyle, this.PlaceStyle);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutPlowCode), SrcConstants.PageLayoutPlowCode, this.PlowCode);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutResizePage), SrcConstants.PageLayoutResizePage, this.ResizePage);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutRouteStyle), SrcConstants.PageLayoutRouteStyle, this.RouteStyle);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutShapeSplit), SrcConstants.PageLayoutShapeSplit, this.ShapeSplit);
        }
    }
}