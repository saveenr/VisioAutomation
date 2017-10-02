using System.Collections.Generic;
using VisioAutomation.ShapeSheet;

namespace VisioPowerShell.Models
{
    public class PageCells : VisioPowerShell.Models.BaseCells
    {

        // Page Format
        public string PageDrawingScale;
        public string PageDrawingScaleType;
        public string PageDrawingSizeType;
        public string PageHeight;
        public string PageScale;
        public string PageWidth;
        public string PageShadowObliqueAngle;
        public string PageShadowOffsetX;
        public string PageShadowOffsetY;
        public string PageShadowScaleFactor;
        public string PageShadowType;
        public string UIVisibility;
        public string PageDrawingResizeType;
        public string PageInhibitSnap;

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

        // Page Print
        public string PrintCenterX;
        public string PrintCenterY;
        public string PrintGrid;
        public string PrintLeftMargin;
        public string PrintPageOrientation;
        public string PrintPaperKind;
        public string PrintPaperSource;
        public string PrintRightMargin;
        public string PrintScaleX;
        public string PrintScaleY;
        public string PrintTopMargin;
        public string PrintBottomMargin;

        public string XGridDensity;
        public string XGridOrigin;
        public string XGridSpacing;
        public string XRulerDensity;
        public string XRulerOrigin;
        public string YGridDensity;
        public string YGridOrigin;
        public string YGridSpacing;
        public string YRulerDensity;
        public string YRulerOrigin;

        public override IEnumerable<CellTuple> GetCellTuples()
        {
            yield return new CellTuple(nameof(SrcConstants.PageDrawingResizeType), SrcConstants.PageDrawingResizeType, this.PageDrawingResizeType);
            yield return new CellTuple(nameof(SrcConstants.PageDrawingScale), SrcConstants.PageDrawingScale, this.PageDrawingScale);
            yield return new CellTuple(nameof(SrcConstants.PageDrawingScaleType), SrcConstants.PageDrawingScaleType, this.PageDrawingScaleType);
            yield return new CellTuple(nameof(SrcConstants.PageDrawingSizeType), SrcConstants.PageDrawingSizeType, this.PageDrawingSizeType);
            yield return new CellTuple(nameof(SrcConstants.PageHeight), SrcConstants.PageHeight, this.PageHeight);
            yield return new CellTuple(nameof(SrcConstants.PageWidth), SrcConstants.PageWidth, this.PageWidth);
            yield return new CellTuple(nameof(SrcConstants.PageScale), SrcConstants.PageScale, this.PageScale);
            yield return new CellTuple(nameof(SrcConstants.PageShadowType), SrcConstants.PageShadowType, this.PageShadowType);
            yield return new CellTuple(nameof(SrcConstants.PageShadowObliqueAngle), SrcConstants.PageShadowObliqueAngle, this.PageShadowObliqueAngle);
            yield return new CellTuple(nameof(SrcConstants.PageShadowOffsetX), SrcConstants.PageShadowOffsetX, this.PageShadowOffsetX);
            yield return new CellTuple(nameof(SrcConstants.PageShadowOffsetY), SrcConstants.PageShadowOffsetY, this.PageShadowOffsetY);
            yield return new CellTuple(nameof(SrcConstants.PageShadowScaleFactor), SrcConstants.PageShadowScaleFactor, this.PageShadowScaleFactor);
            yield return new CellTuple(nameof(SrcConstants.PageInhibitSnap), SrcConstants.PageInhibitSnap, this.PageInhibitSnap);

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

            yield return new CellTuple(nameof(SrcConstants.PrintCenterX), SrcConstants.PrintCenterX, this.PrintCenterX);
            yield return new CellTuple(nameof(SrcConstants.PrintCenterY), SrcConstants.PrintCenterY, this.PrintCenterY);
            yield return new CellTuple(nameof(SrcConstants.PrintGrid), SrcConstants.PrintGrid, this.PrintGrid);
            yield return new CellTuple(nameof(SrcConstants.PrintLeftMargin), SrcConstants.PrintLeftMargin, this.PrintLeftMargin);
            yield return new CellTuple(nameof(SrcConstants.PrintPageOrientation), SrcConstants.PrintPageOrientation, this.PrintPageOrientation);
            yield return new CellTuple(nameof(SrcConstants.PrintPaperKind), SrcConstants.PrintPaperKind, this.PrintPaperKind);
            yield return new CellTuple(nameof(SrcConstants.PrintPaperSource), SrcConstants.PrintPaperSource, this.PrintPaperSource);
            yield return new CellTuple(nameof(SrcConstants.PrintRightMargin), SrcConstants.PrintRightMargin, this.PrintRightMargin);
            yield return new CellTuple(nameof(SrcConstants.PrintScaleX), SrcConstants.PrintScaleX, this.PrintScaleX);
            yield return new CellTuple(nameof(SrcConstants.PrintScaleY), SrcConstants.PrintScaleY, this.PrintScaleY);
            yield return new CellTuple(nameof(SrcConstants.PrintTopMargin), SrcConstants.PrintTopMargin, this.PrintTopMargin);
            yield return new CellTuple(nameof(SrcConstants.PrintBottomMargin), SrcConstants.PrintBottomMargin, this.PrintBottomMargin);

            yield return new CellTuple(nameof(SrcConstants.XGridDensity), SrcConstants.XGridDensity, this.XGridDensity);
            yield return new CellTuple(nameof(SrcConstants.XGridOrigin), SrcConstants.XGridOrigin, this.XGridOrigin);
            yield return new CellTuple(nameof(SrcConstants.XGridSpacing), SrcConstants.XGridSpacing, this.XGridSpacing);
            yield return new CellTuple(nameof(SrcConstants.XRulerDensity), SrcConstants.XRulerDensity, this.XRulerDensity);
            yield return new CellTuple(nameof(SrcConstants.XRulerOrigin), SrcConstants.XRulerOrigin, this.XRulerOrigin);
            yield return new CellTuple(nameof(SrcConstants.YGridDensity), SrcConstants.YGridDensity, this.YGridDensity);
            yield return new CellTuple(nameof(SrcConstants.YGridOrigin), SrcConstants.YGridOrigin, this.YGridOrigin);
            yield return new CellTuple(nameof(SrcConstants.YGridSpacing), SrcConstants.YGridSpacing, this.YGridSpacing);
            yield return new CellTuple(nameof(SrcConstants.YRulerDensity), SrcConstants.YRulerDensity, this.YRulerDensity);
            yield return new CellTuple(nameof(SrcConstants.YRulerOrigin), SrcConstants.YRulerOrigin, this.YRulerOrigin);
        }
    }
}