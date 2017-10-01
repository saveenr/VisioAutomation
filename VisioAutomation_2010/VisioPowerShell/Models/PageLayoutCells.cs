using System.Collections.Generic;
using VisioAutomation.ShapeSheet;

namespace VisioPowerShell.Models
{
    public class PageLayoutCells : VisioPowerShell.Models.BaseCells
    {

        // Page Layout
        public string PageLayoutAvenueSizeX;
        public string PageLayoutAvenueSizeY;
        public string PageLayoutAvoidPageBreaks;
        public string PageLayoutBlockSizeX;
        public string PageLayoutBlockSizeY;
        public string PageLayoutControlAsInput;
        public string PageLayoutDynamicsOff;
        public string PageLayoutEnableGrid;
        public string PageLayoutLineAdjustFrom;
        public string PageLayoutLineAdjustTo;
        public string PageLayoutLineJumpCode;
        public string PageLayoutLineJumpDirX;
        public string PageLayoutLineJumpDirY;
        public string PageLayoutLineJumpFactorX;
        public string PageLayoutLineJumpFactorY;
        public string PageLayoutLineJumpStyle;
        public string PageLayoutLineRouteExt;
        public string PageLayoutLineToLineX;
        public string PageLayoutLineToLineY;
        public string PageLayoutLineToNodeX;
        public string PageLayoutLineToNodeY;
        public string PageLayoutPlaceDepth;
        public string PageLayoutPlaceFlip;
        public string PageLayoutPlaceStyle;
        public string PageLayoutPlowCode;
        public string PageLayoutResizePage;
        public string PageLayoutRouteStyle;
        public string PageLayoutShapeSplit;

        public override IEnumerable<CellTuple> GetCellTuples()
        {
            yield return new CellTuple(nameof(SrcConstants.PageLayoutAvenueSizeX), SrcConstants.PageLayoutAvenueSizeX, this.PageLayoutAvenueSizeX);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutAvenueSizeY), SrcConstants.PageLayoutAvenueSizeY, this.PageLayoutAvenueSizeY);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutAvoidPageBreaks), SrcConstants.PageLayoutAvoidPageBreaks, this.PageLayoutAvoidPageBreaks);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutBlockSizeX), SrcConstants.PageLayoutBlockSizeX, this.PageLayoutBlockSizeX);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutBlockSizeY), SrcConstants.PageLayoutBlockSizeY, this.PageLayoutBlockSizeY);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutControlAsInput), SrcConstants.PageLayoutControlAsInput, this.PageLayoutControlAsInput);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutDynamicsOff), SrcConstants.PageLayoutDynamicsOff, this.PageLayoutDynamicsOff);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutEnableGrid), SrcConstants.PageLayoutEnableGrid, this.PageLayoutEnableGrid);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineAdjustFrom), SrcConstants.PageLayoutLineAdjustFrom, this.PageLayoutLineAdjustFrom);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineAdjustTo), SrcConstants.PageLayoutLineAdjustTo, this.PageLayoutLineAdjustTo);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineJumpCode), SrcConstants.PageLayoutLineJumpCode, this.PageLayoutLineJumpCode);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineJumpDirX), SrcConstants.PageLayoutLineJumpDirX, this.PageLayoutLineJumpDirX);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineJumpDirY), SrcConstants.PageLayoutLineJumpDirY, this.PageLayoutLineJumpDirY);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineJumpFactorX), SrcConstants.PageLayoutLineJumpFactorX, this.PageLayoutLineJumpFactorX);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineJumpFactorY), SrcConstants.PageLayoutLineJumpFactorY, this.PageLayoutLineJumpFactorY);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineJumpStyle), SrcConstants.PageLayoutLineJumpStyle, this.PageLayoutLineJumpStyle);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineRouteExt), SrcConstants.PageLayoutLineRouteExt, this.PageLayoutLineRouteExt);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineToLineX), SrcConstants.PageLayoutLineToLineX, this.PageLayoutLineToLineX);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineToLineY), SrcConstants.PageLayoutLineToLineY, this.PageLayoutLineToLineY);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineToNodeX), SrcConstants.PageLayoutLineToNodeX, this.PageLayoutLineToNodeX);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutLineToNodeY), SrcConstants.PageLayoutLineToNodeY, this.PageLayoutLineToNodeY);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutPlaceDepth), SrcConstants.PageLayoutPlaceDepth, this.PageLayoutPlaceDepth);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutPlaceFlip), SrcConstants.PageLayoutPlaceFlip, this.PageLayoutPlaceFlip);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutPlaceStyle), SrcConstants.PageLayoutPlaceStyle, this.PageLayoutPlaceStyle);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutPlowCode), SrcConstants.PageLayoutPlowCode, this.PageLayoutPlowCode);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutResizePage), SrcConstants.PageLayoutResizePage, this.PageLayoutResizePage);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutRouteStyle), SrcConstants.PageLayoutRouteStyle, this.PageLayoutRouteStyle);
            yield return new CellTuple(nameof(SrcConstants.PageLayoutShapeSplit), SrcConstants.PageLayoutShapeSplit, this.PageLayoutShapeSplit);
        }
    }
}