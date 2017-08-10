using System.Collections.Generic;
using VisioAutomation.ShapeSheet;

namespace VisioScripting.Models
{
    /*
     * 
 ['PrintBottomMargin',
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
 'PageDrawingResizeType']
 
         
         
 ['XFormAngle',
 'OneDBeginX',
 'OneDBeginY',
 'LineBeginArrow',
 'LineBeginArrowSize',
 'CharCase',
 'CharColor',
 'CharColorTransparency',
 'CharFont',
 'CharFontScale',
 'CharLetterspace',
 'CharSize',
 'CharStyle',
 'OneDEndX',
 'OneDEndY',
 'LineEndArrow',
 'LineEndArrowSize',
 'FillBackground',
 'FillBackgroundTransparency',
 'FillForeground',
 'FillForegroundTransparency',
 'FillPattern',
 'XFormHeight',
 'LineCap',
 'LineColor',
 'LinePattern',
 'LineWeight',
 'LockAspect',
 'LockBegin',
 'LockCalcWH',
 'LockCrop',
 'LockCustomProp',
 'LockDelete',
 'LockEnd',
 'LockFormat',
 'LockFromGroupFormat',
 'LockGroup',
 'LockHeight',
 'LockMoveX',
 'LockMoveY',
 'LockRotate',
 'LockSelect',
 'LockTextEdit',
 'LockThemeColors',
 'LockThemeEffects',
 'LockVertexEdit',
 'LockWidth',
 'XFormLocPinX',
 'XFormLocPinY',
 'XFormPinX',
 'XFormPinY',
 'LineRounding',
 'GroupSelectMode',
 'FillShadowBackground',
 'FillShadowBackgroundTransparency',
 'FillShadowForeground',
 'FillShadowForegroundTransparency',
 'PageShadowObliqueAngle',
 'PageShadowOffsetX',
 'PageShadowOffsetY',
 'FillShadowPattern',
 'PageShadowScaleFactor',
 'PageShadowType',
 'TextXFormAngle',
 'TextXFormHeight',
 'TextXFormLocPinX',
 'TextXFormLocPinY',
 'TextXFormPinX',
 'TextXFormPinY',
 'TextXFormWidth',
 'XFormWidth']




         */

    public class PageCells
    {
        public string PrintBottomMargin;
        public string PageHeight;
        public string PrintLeftMargin;
        public string PageLayoutLineJumpDirX;
        public string PageLayoutLineJumpDirY;
        public string PrintRightMargin;
        public string PageScale;
        public string PageLayoutShapeSplit;
        public string PrintTopMargin;
        public string PageWidth;
        public string PrintCenterX;
        public string PrintCenterY;
        public string PrintPaperKind;
        public string PrintGrid;
        public string PrintPageOrientation;
        public string PrintScaleX;
        public string PrintScaleY;
        public string PrintPaperSource;
        public string PageDrawingScale;
        public string PageDrawingScaleType;
        public string PageDrawingSizeType;
        public string PageInhibitSnap;
        public string PageShadowObliqueAngle;
        public string PageShadowOffsetX;
        public string PageShadowOffsetY;
        public string PageShadowScaleFactor;
        public string PageShadowType;
        public string PageUIVisibility;
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
        public string PageLayoutAvenueSizeX;
        public string PageLayoutAvenueSizeY;
        public string PageLayoutBlockSizeX;
        public string PageLayoutBlockSizeY;
        public string PageLayoutControlAsInput;
        public string PageLayoutDynamicsOff;
        public string PageLayoutEnableGrid;
        public string PageLayoutLineAdjustFrom;
        public string PageLayoutLineAdjustTo;
        public string PageLayoutLineJumpCode;
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
        public string PageLayoutAvoidPageBreaks;
        public string PageDrawingResizeType;

        public IEnumerable<VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair> GetSrcFormulaPairs()
        {
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PrintBottomMargin, this.PrintBottomMargin);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageHeight, this.PageHeight);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PrintLeftMargin, this.PrintLeftMargin);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutLineJumpDirX, this.PageLayoutLineJumpDirX);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutLineJumpDirY, this.PageLayoutLineJumpDirY);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PrintRightMargin, this.PrintRightMargin);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageScale, this.PageScale);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutShapeSplit, this.PageLayoutShapeSplit);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PrintTopMargin, this.PrintTopMargin);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageWidth, this.PageWidth);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PrintCenterX, this.PrintCenterX);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PrintCenterY, this.PrintCenterY);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PrintPaperKind, this.PrintPaperKind);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PrintGrid, this.PrintGrid);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PrintPageOrientation, this.PrintPageOrientation);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PrintScaleX, this.PrintScaleX);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PrintScaleY, this.PrintScaleY);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PrintPaperSource, this.PrintPaperSource);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageDrawingScale, this.PageDrawingScale);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageDrawingScaleType, this.PageDrawingScaleType);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageDrawingSizeType, this.PageDrawingSizeType);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageInhibitSnap, this.PageInhibitSnap);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageShadowObliqueAngle, this.PageShadowObliqueAngle);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageShadowOffsetX, this.PageShadowOffsetX);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageShadowOffsetY, this.PageShadowOffsetY);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageShadowScaleFactor, this.PageShadowScaleFactor);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageShadowType, this.PageShadowType);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageUIVisibility, this.PageUIVisibility);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.XGridDensity, this.XGridDensity);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.XGridOrigin, this.XGridOrigin);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.XGridSpacing, this.XGridSpacing);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.XRulerDensity, this.XRulerDensity);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.XRulerOrigin, this.XRulerOrigin);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.YGridDensity, this.YGridDensity);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.YGridOrigin, this.YGridOrigin);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.YGridSpacing, this.YGridSpacing);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.YRulerDensity, this.YRulerDensity);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.YRulerOrigin, this.YRulerOrigin);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutAvenueSizeX, this.PageLayoutAvenueSizeX);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutAvenueSizeY, this.PageLayoutAvenueSizeY);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutBlockSizeX, this.PageLayoutBlockSizeX);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutBlockSizeY, this.PageLayoutBlockSizeY);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutControlAsInput, this.PageLayoutControlAsInput);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutDynamicsOff, this.PageLayoutDynamicsOff);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutEnableGrid, this.PageLayoutEnableGrid);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutLineAdjustFrom, this.PageLayoutLineAdjustFrom);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutLineAdjustTo, this.PageLayoutLineAdjustTo);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutLineJumpCode, this.PageLayoutLineJumpCode);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutLineJumpFactorX, this.PageLayoutLineJumpFactorX);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutLineJumpFactorY, this.PageLayoutLineJumpFactorY);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutLineJumpStyle, this.PageLayoutLineJumpStyle);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutLineRouteExt, this.PageLayoutLineRouteExt);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutLineToLineX, this.PageLayoutLineToLineX);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutLineToLineY, this.PageLayoutLineToLineY);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutLineToNodeX, this.PageLayoutLineToNodeX);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutLineToNodeY, this.PageLayoutLineToNodeY);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutPlaceDepth, this.PageLayoutPlaceDepth);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutPlaceFlip, this.PageLayoutPlaceFlip);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutPlaceStyle, this.PageLayoutPlaceStyle);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutPlowCode, this.PageLayoutPlowCode);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutResizePage, this.PageLayoutResizePage);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutRouteStyle, this.PageLayoutRouteStyle);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageLayoutAvoidPageBreaks, this.PageLayoutAvoidPageBreaks);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageDrawingResizeType, this.PageDrawingResizeType);
        }      
    }

    public class ShapeCells
    {
        public string XFormAngle;
        public string OneDBeginX;
        public string OneDBeginY;
        public string LineBeginArrow;
        public string LineBeginArrowSize;
        public string CharCase;
        public string CharColor;
        public string CharColorTransparency;
        public string CharFont;
        public string CharFontScale;
        public string CharLetterspace;
        public string CharSize;
        public string CharStyle;
        public string OneDEndX;
        public string OneDEndY;
        public string LineEndArrow;
        public string LineEndArrowSize;
        public string FillBackground;
        public string FillBackgroundTransparency;
        public string FillForeground;
        public string FillForegroundTransparency;
        public string FillPattern;
        public string XFormHeight;
        public string LineCap;
        public string LineColor;
        public string LinePattern;
        public string LineWeight;
        public string LockAspect;
        public string LockBegin;
        public string LockCalcWH;
        public string LockCrop;
        public string LockCustomProp;
        public string LockDelete;
        public string LockEnd;
        public string LockFormat;
        public string LockFromGroupFormat;
        public string LockGroup;
        public string LockHeight;
        public string LockMoveX;
        public string LockMoveY;
        public string LockRotate;
        public string LockSelect;
        public string LockTextEdit;
        public string LockThemeColors;
        public string LockThemeEffects;
        public string LockVertexEdit;
        public string LockWidth;
        public string XFormLocPinX;
        public string XFormLocPinY;
        public string XFormPinX;
        public string XFormPinY;
        public string LineRounding;
        public string GroupSelectMode;
        public string FillShadowBackground;
        public string FillShadowBackgroundTransparency;
        public string FillShadowForeground;
        public string FillShadowForegroundTransparency;
        public string PageShadowObliqueAngle;
        public string PageShadowOffsetX;
        public string PageShadowOffsetY;
        public string FillShadowPattern;
        public string PageShadowScaleFactor;
        public string PageShadowType;
        public string TextXFormAngle;
        public string TextXFormHeight;
        public string TextXFormLocPinX;
        public string TextXFormLocPinY;
        public string TextXFormPinX;
        public string TextXFormPinY;
        public string TextXFormWidth;
        public string XFormWidth;

        public IEnumerable<VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair> GetSrcFormulaPairs()
        {
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.XFormAngle, this.XFormAngle);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.OneDBeginX, this.OneDBeginX);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.OneDBeginY, this.OneDBeginY);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LineBeginArrow, this.LineBeginArrow);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LineBeginArrowSize, this.LineBeginArrowSize);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.CharCase, this.CharCase);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.CharColor, this.CharColor);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.CharColorTransparency, this.CharColorTransparency);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.CharFont, this.CharFont);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.CharFontScale, this.CharFontScale);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.CharLetterspace, this.CharLetterspace);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.CharSize, this.CharSize);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.CharStyle, this.CharStyle);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.OneDEndX, this.OneDEndX);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.OneDEndY, this.OneDEndY);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LineEndArrow, this.LineEndArrow);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LineEndArrowSize, this.LineEndArrowSize);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.FillBackground, this.FillBackground);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.FillBackgroundTransparency, this.FillBackgroundTransparency);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.FillForeground, this.FillForeground);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.FillForegroundTransparency, this.FillForegroundTransparency);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.FillPattern, this.FillPattern);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.XFormHeight, this.XFormHeight);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LineCap, this.LineCap);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LineColor, this.LineColor);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LinePattern, this.LinePattern);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LineWeight, this.LineWeight);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockAspect, this.LockAspect);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockBegin, this.LockBegin);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockCalcWH, this.LockCalcWH);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockCrop, this.LockCrop);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockCustomProp, this.LockCustomProp);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockDelete, this.LockDelete);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockEnd, this.LockEnd);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockFormat, this.LockFormat);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockFromGroupFormat, this.LockFromGroupFormat);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockGroup, this.LockGroup);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockHeight, this.LockHeight);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockMoveX, this.LockMoveX);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockMoveY, this.LockMoveY);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockRotate, this.LockRotate);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockSelect, this.LockSelect);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockTextEdit, this.LockTextEdit);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockThemeColors, this.LockThemeColors);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockThemeEffects, this.LockThemeEffects);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockVertexEdit, this.LockVertexEdit);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LockWidth, this.LockWidth);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.XFormLocPinX, this.XFormLocPinX);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.XFormLocPinY, this.XFormLocPinY);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.XFormPinX, this.XFormPinX);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.XFormPinY, this.XFormPinY);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.LineRounding, this.LineRounding);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.GroupSelectMode, this.GroupSelectMode);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.FillShadowBackground, this.FillShadowBackground);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.FillShadowBackgroundTransparency, this.FillShadowBackgroundTransparency);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.FillShadowForeground, this.FillShadowForeground);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.FillShadowForegroundTransparency, this.FillShadowForegroundTransparency);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageShadowObliqueAngle, this.PageShadowObliqueAngle);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageShadowOffsetX, this.PageShadowOffsetX);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageShadowOffsetY, this.PageShadowOffsetY);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.FillShadowPattern, this.FillShadowPattern);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageShadowScaleFactor, this.PageShadowScaleFactor);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.PageShadowType, this.PageShadowType);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.TextXFormAngle, this.TextXFormAngle);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.TextXFormHeight, this.TextXFormHeight);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.TextXFormLocPinX, this.TextXFormLocPinX);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.TextXFormLocPinY, this.TextXFormLocPinY);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.TextXFormPinX, this.TextXFormPinX);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.TextXFormPinY, this.TextXFormPinY);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.TextXFormWidth, this.TextXFormWidth);
            yield return new VisioAutomation.ShapeSheet.CellGroups.SrcFormulaPair(VisioAutomation.ShapeSheet.SrcConstants.XFormWidth, this.XFormWidth);
        }
    }

    public class CellSrcDictionary : NamedSrcDictionary
    {
        private static CellSrcDictionary shape_cellmap;
        private static CellSrcDictionary page_cellmap;

        public static CellSrcDictionary GetCellMapForShapes()
        {
            if (CellSrcDictionary.shape_cellmap == null)
            {
                CellSrcDictionary.shape_cellmap = new CellSrcDictionary();
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.XFormAngle)] = SrcConstants.XFormAngle;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.OneDBeginX)] = SrcConstants.OneDBeginX;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.OneDBeginY)] = SrcConstants.OneDBeginY;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LineBeginArrow)] = SrcConstants.LineBeginArrow;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LineBeginArrowSize)] = SrcConstants.LineBeginArrowSize;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.CharCase)] = SrcConstants.CharCase;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.CharColor)] = SrcConstants.CharColor;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.CharColorTransparency)] = SrcConstants.CharColorTransparency;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.CharFont)] = SrcConstants.CharFont;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.CharFontScale)] = SrcConstants.CharFontScale;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.CharLetterspace)] = SrcConstants.CharLetterspace;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.CharSize)] = SrcConstants.CharSize;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.CharStyle)] = SrcConstants.CharStyle;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.OneDEndX)] = SrcConstants.OneDEndX;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.OneDEndY)] = SrcConstants.OneDEndY;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LineEndArrow)] = SrcConstants.LineEndArrow;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LineEndArrowSize)] = SrcConstants.LineEndArrowSize;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.FillBackground)] = SrcConstants.FillBackground;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.FillBackgroundTransparency)] = SrcConstants.FillBackgroundTransparency;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.FillForeground)] = SrcConstants.FillForeground;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.FillForegroundTransparency)] = SrcConstants.FillForegroundTransparency;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.FillPattern)] = SrcConstants.FillPattern;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.XFormHeight)] = SrcConstants.XFormHeight;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LineCap)] = SrcConstants.LineCap;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LineColor)] = SrcConstants.LineColor;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LinePattern)] = SrcConstants.LinePattern;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LineWeight)] = SrcConstants.LineWeight;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockAspect)] = SrcConstants.LockAspect;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockBegin)] = SrcConstants.LockBegin;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockCalcWH)] = SrcConstants.LockCalcWH;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockCrop)] = SrcConstants.LockCrop;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockCustomProp)] = SrcConstants.LockCustomProp;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockDelete)] = SrcConstants.LockDelete;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockEnd)] = SrcConstants.LockEnd;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockFormat)] = SrcConstants.LockFormat;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockFromGroupFormat)] = SrcConstants.LockFromGroupFormat;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockGroup)] = SrcConstants.LockGroup;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockHeight)] = SrcConstants.LockHeight;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockMoveX)] = SrcConstants.LockMoveX;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockMoveY)] = SrcConstants.LockMoveY;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockRotate)] = SrcConstants.LockRotate;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockSelect)] = SrcConstants.LockSelect;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockTextEdit)] = SrcConstants.LockTextEdit;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockThemeColors)] = SrcConstants.LockThemeColors;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockThemeEffects)] = SrcConstants.LockThemeEffects;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockVertexEdit)] = SrcConstants.LockVertexEdit;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockWidth)] = SrcConstants.LockWidth;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.XFormLocPinX)] = SrcConstants.XFormLocPinX;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.XFormLocPinY)] = SrcConstants.XFormLocPinY;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.XFormPinX)] = SrcConstants.XFormPinX;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.XFormPinY)] = SrcConstants.XFormPinY;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LineRounding)] = SrcConstants.LineRounding;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.GroupSelectMode)] = SrcConstants.GroupSelectMode;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.FillShadowBackground)] = SrcConstants.FillShadowBackground;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.FillShadowBackgroundTransparency)] = SrcConstants.FillShadowBackgroundTransparency;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.FillShadowForeground)] = SrcConstants.FillShadowForeground;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.FillShadowForegroundTransparency)] = SrcConstants.FillShadowForegroundTransparency;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.PageShadowObliqueAngle)] = SrcConstants.PageShadowObliqueAngle;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.PageShadowOffsetX)] = SrcConstants.PageShadowOffsetX;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.PageShadowOffsetY)] = SrcConstants.PageShadowOffsetY;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.FillShadowPattern)] = SrcConstants.FillShadowPattern;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.PageShadowScaleFactor)] = SrcConstants.PageShadowScaleFactor;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.PageShadowType)] = SrcConstants.PageShadowType;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.TextXFormAngle)] = SrcConstants.TextXFormAngle;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.TextXFormHeight)] = SrcConstants.TextXFormHeight;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.TextXFormLocPinX)] = SrcConstants.TextXFormLocPinX;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.TextXFormLocPinY)] = SrcConstants.TextXFormLocPinY;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.TextXFormPinX)] = SrcConstants.TextXFormPinX;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.TextXFormPinY)] = SrcConstants.TextXFormPinY;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.TextXFormWidth)] = SrcConstants.TextXFormWidth;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.XFormWidth)] = SrcConstants.XFormWidth;

            }
            return CellSrcDictionary.shape_cellmap;
        }

        public static CellSrcDictionary GetCellMapForPages()
        {
            if (CellSrcDictionary.page_cellmap == null)
            {
                CellSrcDictionary.page_cellmap = new CellSrcDictionary();
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PrintBottomMargin)] = SrcConstants.PrintBottomMargin;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageHeight)] = SrcConstants.PageHeight;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PrintLeftMargin)] = SrcConstants.PrintLeftMargin;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutLineJumpDirX)] = SrcConstants.PageLayoutLineJumpDirX;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutLineJumpDirY)] = SrcConstants.PageLayoutLineJumpDirY;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PrintRightMargin)] = SrcConstants.PrintRightMargin;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageScale)] = SrcConstants.PageScale;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutShapeSplit)] = SrcConstants.PageLayoutShapeSplit;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PrintTopMargin)] = SrcConstants.PrintTopMargin;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageWidth)] = SrcConstants.PageWidth;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PrintCenterX)] = SrcConstants.PrintCenterX;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PrintCenterY)] = SrcConstants.PrintCenterY;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PrintPaperKind)] = SrcConstants.PrintPaperKind;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PrintGrid)] = SrcConstants.PrintGrid;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PrintPageOrientation)] = SrcConstants.PrintPageOrientation;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PrintScaleX)] = SrcConstants.PrintScaleX;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PrintScaleY)] = SrcConstants.PrintScaleY;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PrintPaperSource)] = SrcConstants.PrintPaperSource;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageDrawingScale)] = SrcConstants.PageDrawingScale;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageDrawingScaleType)] = SrcConstants.PageDrawingScaleType;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageDrawingSizeType)] = SrcConstants.PageDrawingSizeType;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageInhibitSnap)] = SrcConstants.PageInhibitSnap;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageShadowObliqueAngle)] = SrcConstants.PageShadowObliqueAngle;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageShadowOffsetX)] = SrcConstants.PageShadowOffsetX;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageShadowOffsetY)] = SrcConstants.PageShadowOffsetY;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageShadowScaleFactor)] = SrcConstants.PageShadowScaleFactor;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageShadowType)] = SrcConstants.PageShadowType;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageUIVisibility)] = SrcConstants.PageUIVisibility;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.XGridDensity)] = SrcConstants.XGridDensity;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.XGridOrigin)] = SrcConstants.XGridOrigin;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.XGridSpacing)] = SrcConstants.XGridSpacing;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.XRulerDensity)] = SrcConstants.XRulerDensity;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.XRulerOrigin)] = SrcConstants.XRulerOrigin;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.YGridDensity)] = SrcConstants.YGridDensity;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.YGridOrigin)] = SrcConstants.YGridOrigin;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.YGridSpacing)] = SrcConstants.YGridSpacing;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.YRulerDensity)] = SrcConstants.YRulerDensity;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.YRulerOrigin)] = SrcConstants.YRulerOrigin;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutAvenueSizeX)] = SrcConstants.PageLayoutAvenueSizeX;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutAvenueSizeY)] = SrcConstants.PageLayoutAvenueSizeY;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutBlockSizeX)] = SrcConstants.PageLayoutBlockSizeX;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutBlockSizeY)] = SrcConstants.PageLayoutBlockSizeY;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutControlAsInput)] = SrcConstants.PageLayoutControlAsInput;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutDynamicsOff)] = SrcConstants.PageLayoutDynamicsOff;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutEnableGrid)] = SrcConstants.PageLayoutEnableGrid;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutLineAdjustFrom)] = SrcConstants.PageLayoutLineAdjustFrom;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutLineAdjustTo)] = SrcConstants.PageLayoutLineAdjustTo;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutLineJumpCode)] = SrcConstants.PageLayoutLineJumpCode;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutLineJumpFactorX)] = SrcConstants.PageLayoutLineJumpFactorX;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutLineJumpFactorY)] = SrcConstants.PageLayoutLineJumpFactorY;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutLineJumpStyle)] = SrcConstants.PageLayoutLineJumpStyle;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutLineRouteExt)] = SrcConstants.PageLayoutLineRouteExt;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutLineToLineX)] = SrcConstants.PageLayoutLineToLineX;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutLineToLineY)] = SrcConstants.PageLayoutLineToLineY;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutLineToNodeX)] = SrcConstants.PageLayoutLineToNodeX;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutLineToNodeY)] = SrcConstants.PageLayoutLineToNodeY;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutPlaceDepth)] = SrcConstants.PageLayoutPlaceDepth;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutPlaceFlip)] = SrcConstants.PageLayoutPlaceFlip;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutPlaceStyle)] = SrcConstants.PageLayoutPlaceStyle;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutPlowCode)] = SrcConstants.PageLayoutPlowCode;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutResizePage)] = SrcConstants.PageLayoutResizePage;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutRouteStyle)] = SrcConstants.PageLayoutRouteStyle;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLayoutAvoidPageBreaks)] = SrcConstants.PageLayoutAvoidPageBreaks;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageDrawingResizeType)] = SrcConstants.PageDrawingResizeType;
            }
            return CellSrcDictionary.page_cellmap;
        }
    }
}

