using System.Collections.Generic;
using SRCCON = VisioAutomation.ShapeSheet.SrcConstants;

namespace VisioPowerShell.Models;

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

    // Grid & Ruler
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

    internal override IEnumerable<Internal.CellTuple> EnumCellTuples()
    {
        yield return new Internal.CellTuple(nameof(SRCCON.PageDrawingResizeType), SRCCON.PageDrawingResizeType, this.PageDrawingResizeType);
        yield return new Internal.CellTuple(nameof(SRCCON.PageDrawingScale), SRCCON.PageDrawingScale, this.PageDrawingScale);
        yield return new Internal.CellTuple(nameof(SRCCON.PageDrawingScaleType), SRCCON.PageDrawingScaleType, this.PageDrawingScaleType);
        yield return new Internal.CellTuple(nameof(SRCCON.PageDrawingSizeType), SRCCON.PageDrawingSizeType, this.PageDrawingSizeType);
        yield return new Internal.CellTuple(nameof(SRCCON.PageHeight), SRCCON.PageHeight, this.PageHeight);
        yield return new Internal.CellTuple(nameof(SRCCON.PageWidth), SRCCON.PageWidth, this.PageWidth);
        yield return new Internal.CellTuple(nameof(SRCCON.PageScale), SRCCON.PageScale, this.PageScale);
        yield return new Internal.CellTuple(nameof(SRCCON.PageShadowType), SRCCON.PageShadowType, this.PageShadowType);
        yield return new Internal.CellTuple(nameof(SRCCON.PageShadowObliqueAngle), SRCCON.PageShadowObliqueAngle, this.PageShadowObliqueAngle);
        yield return new Internal.CellTuple(nameof(SRCCON.PageShadowOffsetX), SRCCON.PageShadowOffsetX, this.PageShadowOffsetX);
        yield return new Internal.CellTuple(nameof(SRCCON.PageShadowOffsetY), SRCCON.PageShadowOffsetY, this.PageShadowOffsetY);
        yield return new Internal.CellTuple(nameof(SRCCON.PageShadowScaleFactor), SRCCON.PageShadowScaleFactor, this.PageShadowScaleFactor);
        yield return new Internal.CellTuple(nameof(SRCCON.PageInhibitSnap), SRCCON.PageInhibitSnap, this.PageInhibitSnap);

        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutAvenueSizeX), SRCCON.PageLayoutAvenueSizeX, this.PageLayoutAvenueSizeX);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutAvenueSizeY), SRCCON.PageLayoutAvenueSizeY, this.PageLayoutAvenueSizeY);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutAvoidPageBreaks), SRCCON.PageLayoutAvoidPageBreaks, this.PageLayoutAvoidPageBreaks);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutBlockSizeX), SRCCON.PageLayoutBlockSizeX, this.PageLayoutBlockSizeX);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutBlockSizeY), SRCCON.PageLayoutBlockSizeY, this.PageLayoutBlockSizeY);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutControlAsInput), SRCCON.PageLayoutControlAsInput, this.PageLayoutControlAsInput);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutDynamicsOff), SRCCON.PageLayoutDynamicsOff, this.PageLayoutDynamicsOff);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutEnableGrid), SRCCON.PageLayoutEnableGrid, this.PageLayoutEnableGrid);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutLineAdjustFrom), SRCCON.PageLayoutLineAdjustFrom, this.PageLayoutLineAdjustFrom);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutLineAdjustTo), SRCCON.PageLayoutLineAdjustTo, this.PageLayoutLineAdjustTo);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutLineJumpCode), SRCCON.PageLayoutLineJumpCode, this.PageLayoutLineJumpCode);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutLineJumpDirX), SRCCON.PageLayoutLineJumpDirX, this.PageLayoutLineJumpDirX);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutLineJumpDirY), SRCCON.PageLayoutLineJumpDirY, this.PageLayoutLineJumpDirY);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutLineJumpFactorX), SRCCON.PageLayoutLineJumpFactorX, this.PageLayoutLineJumpFactorX);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutLineJumpFactorY), SRCCON.PageLayoutLineJumpFactorY, this.PageLayoutLineJumpFactorY);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutLineJumpStyle), SRCCON.PageLayoutLineJumpStyle, this.PageLayoutLineJumpStyle);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutLineRouteExt), SRCCON.PageLayoutLineRouteExt, this.PageLayoutLineRouteExt);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutLineToLineX), SRCCON.PageLayoutLineToLineX, this.PageLayoutLineToLineX);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutLineToLineY), SRCCON.PageLayoutLineToLineY, this.PageLayoutLineToLineY);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutLineToNodeX), SRCCON.PageLayoutLineToNodeX, this.PageLayoutLineToNodeX);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutLineToNodeY), SRCCON.PageLayoutLineToNodeY, this.PageLayoutLineToNodeY);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutPlaceDepth), SRCCON.PageLayoutPlaceDepth, this.PageLayoutPlaceDepth);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutPlaceFlip), SRCCON.PageLayoutPlaceFlip, this.PageLayoutPlaceFlip);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutPlaceStyle), SRCCON.PageLayoutPlaceStyle, this.PageLayoutPlaceStyle);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutPlowCode), SRCCON.PageLayoutPlowCode, this.PageLayoutPlowCode);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutResizePage), SRCCON.PageLayoutResizePage, this.PageLayoutResizePage);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutRouteStyle), SRCCON.PageLayoutRouteStyle, this.PageLayoutRouteStyle);
        yield return new Internal.CellTuple(nameof(SRCCON.PageLayoutShapeSplit), SRCCON.PageLayoutShapeSplit, this.PageLayoutShapeSplit);

        yield return new Internal.CellTuple(nameof(SRCCON.PrintCenterX), SRCCON.PrintCenterX, this.PrintCenterX);
        yield return new Internal.CellTuple(nameof(SRCCON.PrintCenterY), SRCCON.PrintCenterY, this.PrintCenterY);
        yield return new Internal.CellTuple(nameof(SRCCON.PrintGrid), SRCCON.PrintGrid, this.PrintGrid);
        yield return new Internal.CellTuple(nameof(SRCCON.PrintLeftMargin), SRCCON.PrintLeftMargin, this.PrintLeftMargin);
        yield return new Internal.CellTuple(nameof(SRCCON.PrintPageOrientation), SRCCON.PrintPageOrientation, this.PrintPageOrientation);
        yield return new Internal.CellTuple(nameof(SRCCON.PrintPaperKind), SRCCON.PrintPaperKind, this.PrintPaperKind);
        yield return new Internal.CellTuple(nameof(SRCCON.PrintPaperSource), SRCCON.PrintPaperSource, this.PrintPaperSource);
        yield return new Internal.CellTuple(nameof(SRCCON.PrintRightMargin), SRCCON.PrintRightMargin, this.PrintRightMargin);
        yield return new Internal.CellTuple(nameof(SRCCON.PrintScaleX), SRCCON.PrintScaleX, this.PrintScaleX);
        yield return new Internal.CellTuple(nameof(SRCCON.PrintScaleY), SRCCON.PrintScaleY, this.PrintScaleY);
        yield return new Internal.CellTuple(nameof(SRCCON.PrintTopMargin), SRCCON.PrintTopMargin, this.PrintTopMargin);
        yield return new Internal.CellTuple(nameof(SRCCON.PrintBottomMargin), SRCCON.PrintBottomMargin, this.PrintBottomMargin);

        yield return new Internal.CellTuple(nameof(SRCCON.XGridDensity), SRCCON.XGridDensity, this.XGridDensity);
        yield return new Internal.CellTuple(nameof(SRCCON.XGridOrigin), SRCCON.XGridOrigin, this.XGridOrigin);
        yield return new Internal.CellTuple(nameof(SRCCON.XGridSpacing), SRCCON.XGridSpacing, this.XGridSpacing);
        yield return new Internal.CellTuple(nameof(SRCCON.XRulerDensity), SRCCON.XRulerDensity, this.XRulerDensity);
        yield return new Internal.CellTuple(nameof(SRCCON.XRulerOrigin), SRCCON.XRulerOrigin, this.XRulerOrigin);
        yield return new Internal.CellTuple(nameof(SRCCON.YGridDensity), SRCCON.YGridDensity, this.YGridDensity);
        yield return new Internal.CellTuple(nameof(SRCCON.YGridOrigin), SRCCON.YGridOrigin, this.YGridOrigin);
        yield return new Internal.CellTuple(nameof(SRCCON.YGridSpacing), SRCCON.YGridSpacing, this.YGridSpacing);
        yield return new Internal.CellTuple(nameof(SRCCON.YRulerDensity), SRCCON.YRulerDensity, this.YRulerDensity);
        yield return new Internal.CellTuple(nameof(SRCCON.YRulerOrigin), SRCCON.YRulerOrigin, this.YRulerOrigin);
    }
}