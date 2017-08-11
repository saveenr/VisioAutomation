using System.Collections.Generic;
using SRCCON = VisioAutomation.ShapeSheet.SrcConstants;
using CellTuple = VisioPowerShell.Models.CellTuple;

namespace VisioPowerShell.Models
{
    public class PageCells : VisioPowerShell.Models.BaseCells
    {
        public string PageDrawingResizeType;
        public string PageDrawingScale;
        public string PageDrawingScaleType;
        public string PageDrawingSizeType;
        public string PageHeight;
        public string PageInhibitSnap;
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
        public string PageScale;
        public string PageShadowObliqueAngle;
        public string PageShadowOffsetX;
        public string PageShadowOffsetY;
        public string PageShadowScaleFactor;
        public string PageShadowType;
        public string PageUIVisibility;
        public string PageWidth;
        public string PrintBottomMargin;
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

        private static VisioPowerShell.Models.NamedSrcDictionary cellmap;

        public static VisioPowerShell.Models.NamedSrcDictionary GetCellDictionary()
        {
            if (cellmap == null)
            {
                var cells = new VisioPowerShell.Models.PageCells();
                cellmap = VisioPowerShell.Models.NamedSrcDictionary.FromCells(cells);
            }
            return cellmap;
        }

        public override IEnumerable<CellTuple> GetCellTuples()
        {
            yield return new CellTuple(nameof(SRCCON.PageDrawingResizeType), SRCCON.PageDrawingResizeType, this.PageDrawingResizeType);
            yield return new CellTuple(nameof(SRCCON.PageDrawingScale), SRCCON.PageDrawingScale, this.PageDrawingScale);
            yield return new CellTuple(nameof(SRCCON.PageDrawingScaleType), SRCCON.PageDrawingScaleType, this.PageDrawingScaleType);
            yield return new CellTuple(nameof(SRCCON.PageDrawingSizeType), SRCCON.PageDrawingSizeType, this.PageDrawingSizeType);
            yield return new CellTuple(nameof(SRCCON.PageHeight), SRCCON.PageHeight, this.PageHeight);
            yield return new CellTuple(nameof(SRCCON.PageInhibitSnap), SRCCON.PageInhibitSnap, this.PageInhibitSnap);
            yield return new CellTuple(nameof(SRCCON.PageLayoutAvenueSizeX), SRCCON.PageLayoutAvenueSizeX, this.PageLayoutAvenueSizeX);
            yield return new CellTuple(nameof(SRCCON.PageLayoutAvenueSizeY), SRCCON.PageLayoutAvenueSizeY, this.PageLayoutAvenueSizeY);
            yield return new CellTuple(nameof(SRCCON.PageLayoutAvoidPageBreaks), SRCCON.PageLayoutAvoidPageBreaks, this.PageLayoutAvoidPageBreaks);
            yield return new CellTuple(nameof(SRCCON.PageLayoutBlockSizeX), SRCCON.PageLayoutBlockSizeX, this.PageLayoutBlockSizeX);
            yield return new CellTuple(nameof(SRCCON.PageLayoutBlockSizeY), SRCCON.PageLayoutBlockSizeY, this.PageLayoutBlockSizeY);
            yield return new CellTuple(nameof(SRCCON.PageLayoutControlAsInput), SRCCON.PageLayoutControlAsInput, this.PageLayoutControlAsInput);
            yield return new CellTuple(nameof(SRCCON.PageLayoutDynamicsOff), SRCCON.PageLayoutDynamicsOff, this.PageLayoutDynamicsOff);
            yield return new CellTuple(nameof(SRCCON.PageLayoutEnableGrid), SRCCON.PageLayoutEnableGrid, this.PageLayoutEnableGrid);
            yield return new CellTuple(nameof(SRCCON.PageLayoutLineAdjustFrom), SRCCON.PageLayoutLineAdjustFrom, this.PageLayoutLineAdjustFrom);
            yield return new CellTuple(nameof(SRCCON.PageLayoutLineAdjustTo), SRCCON.PageLayoutLineAdjustTo, this.PageLayoutLineAdjustTo);
            yield return new CellTuple(nameof(SRCCON.PageLayoutLineJumpCode), SRCCON.PageLayoutLineJumpCode, this.PageLayoutLineJumpCode);
            yield return new CellTuple(nameof(SRCCON.PageLayoutLineJumpDirX), SRCCON.PageLayoutLineJumpDirX, this.PageLayoutLineJumpDirX);
            yield return new CellTuple(nameof(SRCCON.PageLayoutLineJumpDirY), SRCCON.PageLayoutLineJumpDirY, this.PageLayoutLineJumpDirY);
            yield return new CellTuple(nameof(SRCCON.PageLayoutLineJumpFactorX), SRCCON.PageLayoutLineJumpFactorX, this.PageLayoutLineJumpFactorX);
            yield return new CellTuple(nameof(SRCCON.PageLayoutLineJumpFactorY), SRCCON.PageLayoutLineJumpFactorY, this.PageLayoutLineJumpFactorY);
            yield return new CellTuple(nameof(SRCCON.PageLayoutLineJumpStyle), SRCCON.PageLayoutLineJumpStyle, this.PageLayoutLineJumpStyle);
            yield return new CellTuple(nameof(SRCCON.PageLayoutLineRouteExt), SRCCON.PageLayoutLineRouteExt, this.PageLayoutLineRouteExt);
            yield return new CellTuple(nameof(SRCCON.PageLayoutLineToLineX), SRCCON.PageLayoutLineToLineX, this.PageLayoutLineToLineX);
            yield return new CellTuple(nameof(SRCCON.PageLayoutLineToLineY), SRCCON.PageLayoutLineToLineY, this.PageLayoutLineToLineY);
            yield return new CellTuple(nameof(SRCCON.PageLayoutLineToNodeX), SRCCON.PageLayoutLineToNodeX, this.PageLayoutLineToNodeX);
            yield return new CellTuple(nameof(SRCCON.PageLayoutLineToNodeY), SRCCON.PageLayoutLineToNodeY, this.PageLayoutLineToNodeY);
            yield return new CellTuple(nameof(SRCCON.PageLayoutPlaceDepth), SRCCON.PageLayoutPlaceDepth, this.PageLayoutPlaceDepth);
            yield return new CellTuple(nameof(SRCCON.PageLayoutPlaceFlip), SRCCON.PageLayoutPlaceFlip, this.PageLayoutPlaceFlip);
            yield return new CellTuple(nameof(SRCCON.PageLayoutPlaceStyle), SRCCON.PageLayoutPlaceStyle, this.PageLayoutPlaceStyle);
            yield return new CellTuple(nameof(SRCCON.PageLayoutPlowCode), SRCCON.PageLayoutPlowCode, this.PageLayoutPlowCode);
            yield return new CellTuple(nameof(SRCCON.PageLayoutResizePage), SRCCON.PageLayoutResizePage, this.PageLayoutResizePage);
            yield return new CellTuple(nameof(SRCCON.PageLayoutRouteStyle), SRCCON.PageLayoutRouteStyle, this.PageLayoutRouteStyle);
            yield return new CellTuple(nameof(SRCCON.PageLayoutShapeSplit), SRCCON.PageLayoutShapeSplit, this.PageLayoutShapeSplit);
            yield return new CellTuple(nameof(SRCCON.PageScale), SRCCON.PageScale, this.PageScale);
            yield return new CellTuple(nameof(SRCCON.PageShadowObliqueAngle), SRCCON.PageShadowObliqueAngle, this.PageShadowObliqueAngle);
            yield return new CellTuple(nameof(SRCCON.PageShadowOffsetX), SRCCON.PageShadowOffsetX, this.PageShadowOffsetX);
            yield return new CellTuple(nameof(SRCCON.PageShadowOffsetY), SRCCON.PageShadowOffsetY, this.PageShadowOffsetY);
            yield return new CellTuple(nameof(SRCCON.PageShadowScaleFactor), SRCCON.PageShadowScaleFactor, this.PageShadowScaleFactor);
            yield return new CellTuple(nameof(SRCCON.PageShadowType), SRCCON.PageShadowType, this.PageShadowType);
            yield return new CellTuple(nameof(SRCCON.PageUIVisibility), SRCCON.PageUIVisibility, this.PageUIVisibility);
            yield return new CellTuple(nameof(SRCCON.PageWidth), SRCCON.PageWidth, this.PageWidth);
            yield return new CellTuple(nameof(SRCCON.PrintBottomMargin), SRCCON.PrintBottomMargin, this.PrintBottomMargin);
            yield return new CellTuple(nameof(SRCCON.PrintCenterX), SRCCON.PrintCenterX, this.PrintCenterX);
            yield return new CellTuple(nameof(SRCCON.PrintCenterY), SRCCON.PrintCenterY, this.PrintCenterY);
            yield return new CellTuple(nameof(SRCCON.PrintGrid), SRCCON.PrintGrid, this.PrintGrid);
            yield return new CellTuple(nameof(SRCCON.PrintLeftMargin), SRCCON.PrintLeftMargin, this.PrintLeftMargin);
            yield return new CellTuple(nameof(SRCCON.PrintPageOrientation), SRCCON.PrintPageOrientation, this.PrintPageOrientation);
            yield return new CellTuple(nameof(SRCCON.PrintPaperKind), SRCCON.PrintPaperKind, this.PrintPaperKind);
            yield return new CellTuple(nameof(SRCCON.PrintPaperSource), SRCCON.PrintPaperSource, this.PrintPaperSource);
            yield return new CellTuple(nameof(SRCCON.PrintRightMargin), SRCCON.PrintRightMargin, this.PrintRightMargin);
            yield return new CellTuple(nameof(SRCCON.PrintScaleX), SRCCON.PrintScaleX, this.PrintScaleX);
            yield return new CellTuple(nameof(SRCCON.PrintScaleY), SRCCON.PrintScaleY, this.PrintScaleY);
            yield return new CellTuple(nameof(SRCCON.PrintTopMargin), SRCCON.PrintTopMargin, this.PrintTopMargin);
            yield return new CellTuple(nameof(SRCCON.XGridDensity), SRCCON.XGridDensity, this.XGridDensity);
            yield return new CellTuple(nameof(SRCCON.XGridOrigin), SRCCON.XGridOrigin, this.XGridOrigin);
            yield return new CellTuple(nameof(SRCCON.XGridSpacing), SRCCON.XGridSpacing, this.XGridSpacing);
            yield return new CellTuple(nameof(SRCCON.XRulerDensity), SRCCON.XRulerDensity, this.XRulerDensity);
            yield return new CellTuple(nameof(SRCCON.XRulerOrigin), SRCCON.XRulerOrigin, this.XRulerOrigin);
            yield return new CellTuple(nameof(SRCCON.YGridDensity), SRCCON.YGridDensity, this.YGridDensity);
            yield return new CellTuple(nameof(SRCCON.YGridOrigin), SRCCON.YGridOrigin, this.YGridOrigin);
            yield return new CellTuple(nameof(SRCCON.YGridSpacing), SRCCON.YGridSpacing, this.YGridSpacing);
            yield return new CellTuple(nameof(SRCCON.YRulerDensity), SRCCON.YRulerDensity, this.YRulerDensity);
            yield return new CellTuple(nameof(SRCCON.YRulerOrigin), SRCCON.YRulerOrigin, this.YRulerOrigin);
        }

    }
}


/*
Page cells

    [ 'PrintBottomMargin',
    'PageHeight',
    'PrintLeftMargin',
    'PageLayoutLineJumpDirX',
    'PageLayoutLineJumpDirY',
    'PrintRightMargin',
    'PageScale',
    'PageLayoutShapeSplit',
    'PrintTopMargin',
    'PageWidth',
    'PrintCenterX',
    'PrintCenterY',
    'PrintPaperKind',
    'PrintGrid',
    'PrintPageOrientation',
    'PrintScaleX',
    'PrintScaleY',
    'PrintPaperSource',
    'PageDrawingScale',
    'PageDrawingScaleType',
    'PageDrawingSizeType',
    'PageInhibitSnap',
    'PageShadowObliqueAngle',
    'PageShadowOffsetX',
    'PageShadowOffsetY',
    'PageShadowScaleFactor',
    'PageShadowType',
    'PageUIVisibility',
    'XGridDensity',
    'XGridOrigin',
    'XGridSpacing',
    'XRulerDensity',
    'XRulerOrigin',
    'YGridDensity',
    'YGridOrigin',
    'YGridSpacing',
    'YRulerDensity',
    'YRulerOrigin',
    'PageLayoutAvenueSizeX',
    'PageLayoutAvenueSizeY',
    'PageLayoutBlockSizeX',
    'PageLayoutBlockSizeY',
    'PageLayoutControlAsInput',
    'PageLayoutDynamicsOff',
    'PageLayoutEnableGrid',
    'PageLayoutLineAdjustFrom',
    'PageLayoutLineAdjustTo',
    'PageLayoutLineJumpCode',
    'PageLayoutLineJumpFactorX',
    'PageLayoutLineJumpFactorY',
    'PageLayoutLineJumpStyle',
    'PageLayoutLineRouteExt',
    'PageLayoutLineToLineX',
    'PageLayoutLineToLineY',
    'PageLayoutLineToNodeX',
    'PageLayoutLineToNodeY',
    'PageLayoutPlaceDepth',
    'PageLayoutPlaceFlip',
    'PageLayoutPlaceStyle',
    'PageLayoutPlowCode',
    'PageLayoutResizePage',
    'PageLayoutRouteStyle',
    'PageLayoutAvoidPageBreaks',
    'PageDrawingResizeType' ]
    */

