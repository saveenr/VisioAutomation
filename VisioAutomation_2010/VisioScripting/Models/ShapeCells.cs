using System.Collections.Generic;
using SRCCON = VisioAutomation.ShapeSheet.SrcConstants;

namespace VisioScripting.Models
{
    public class ShapeCells : BaseCells
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

        public override IEnumerable<CellTuple> GetSrcFormulaPairs()
        {
            yield return new CellTuple(nameof(SRCCON.CharCase), SRCCON.CharCase, this.CharCase);
            yield return new CellTuple(nameof(SRCCON.CharColor), SRCCON.CharColor, this.CharColor);
            yield return new CellTuple(nameof(SRCCON.CharColorTransparency), SRCCON.CharColorTransparency, this.CharColorTransparency);
            yield return new CellTuple(nameof(SRCCON.CharFont), SRCCON.CharFont, this.CharFont);
            yield return new CellTuple(nameof(SRCCON.CharFontScale), SRCCON.CharFontScale, this.CharFontScale);
            yield return new CellTuple(nameof(SRCCON.CharLetterspace), SRCCON.CharLetterspace, this.CharLetterspace);
            yield return new CellTuple(nameof(SRCCON.CharSize), SRCCON.CharSize, this.CharSize);
            yield return new CellTuple(nameof(SRCCON.CharStyle), SRCCON.CharStyle, this.CharStyle);
            yield return new CellTuple(nameof(SRCCON.FillBackground), SRCCON.FillBackground, this.FillBackground);
            yield return new CellTuple(nameof(SRCCON.FillBackgroundTransparency), SRCCON.FillBackgroundTransparency, this.FillBackgroundTransparency);
            yield return new CellTuple(nameof(SRCCON.FillForeground), SRCCON.FillForeground, this.FillForeground);
            yield return new CellTuple(nameof(SRCCON.FillForegroundTransparency), SRCCON.FillForegroundTransparency, this.FillForegroundTransparency);
            yield return new CellTuple(nameof(SRCCON.FillPattern), SRCCON.FillPattern, this.FillPattern);
            yield return new CellTuple(nameof(SRCCON.FillShadowBackground), SRCCON.FillShadowBackground, this.FillShadowBackground);
            yield return new CellTuple(nameof(SRCCON.FillShadowBackgroundTransparency), SRCCON.FillShadowBackgroundTransparency, this.FillShadowBackgroundTransparency);
            yield return new CellTuple(nameof(SRCCON.FillShadowForeground), SRCCON.FillShadowForeground, this.FillShadowForeground);
            yield return new CellTuple(nameof(SRCCON.FillShadowForegroundTransparency), SRCCON.FillShadowForegroundTransparency, this.FillShadowForegroundTransparency);
            yield return new CellTuple(nameof(SRCCON.FillShadowPattern), SRCCON.FillShadowPattern, this.FillShadowPattern);
            yield return new CellTuple(nameof(SRCCON.GroupSelectMode), SRCCON.GroupSelectMode, this.GroupSelectMode);
            yield return new CellTuple(nameof(SRCCON.LineBeginArrow), SRCCON.LineBeginArrow, this.LineBeginArrow);
            yield return new CellTuple(nameof(SRCCON.LineBeginArrowSize), SRCCON.LineBeginArrowSize, this.LineBeginArrowSize);
            yield return new CellTuple(nameof(SRCCON.LineCap), SRCCON.LineCap, this.LineCap);
            yield return new CellTuple(nameof(SRCCON.LineColor), SRCCON.LineColor, this.LineColor);
            yield return new CellTuple(nameof(SRCCON.LineEndArrow), SRCCON.LineEndArrow, this.LineEndArrow);
            yield return new CellTuple(nameof(SRCCON.LineEndArrowSize), SRCCON.LineEndArrowSize, this.LineEndArrowSize);
            yield return new CellTuple(nameof(SRCCON.LinePattern), SRCCON.LinePattern, this.LinePattern);
            yield return new CellTuple(nameof(SRCCON.LineRounding), SRCCON.LineRounding, this.LineRounding);
            yield return new CellTuple(nameof(SRCCON.LineWeight), SRCCON.LineWeight, this.LineWeight);
            yield return new CellTuple(nameof(SRCCON.LockAspect), SRCCON.LockAspect, this.LockAspect);
            yield return new CellTuple(nameof(SRCCON.LockBegin), SRCCON.LockBegin, this.LockBegin);
            yield return new CellTuple(nameof(SRCCON.LockCalcWH), SRCCON.LockCalcWH, this.LockCalcWH);
            yield return new CellTuple(nameof(SRCCON.LockCrop), SRCCON.LockCrop, this.LockCrop);
            yield return new CellTuple(nameof(SRCCON.LockCustomProp), SRCCON.LockCustomProp, this.LockCustomProp);
            yield return new CellTuple(nameof(SRCCON.LockDelete), SRCCON.LockDelete, this.LockDelete);
            yield return new CellTuple(nameof(SRCCON.LockEnd), SRCCON.LockEnd, this.LockEnd);
            yield return new CellTuple(nameof(SRCCON.LockFormat), SRCCON.LockFormat, this.LockFormat);
            yield return new CellTuple(nameof(SRCCON.LockFromGroupFormat), SRCCON.LockFromGroupFormat, this.LockFromGroupFormat);
            yield return new CellTuple(nameof(SRCCON.LockGroup), SRCCON.LockGroup, this.LockGroup);
            yield return new CellTuple(nameof(SRCCON.LockHeight), SRCCON.LockHeight, this.LockHeight);
            yield return new CellTuple(nameof(SRCCON.LockMoveX), SRCCON.LockMoveX, this.LockMoveX);
            yield return new CellTuple(nameof(SRCCON.LockMoveY), SRCCON.LockMoveY, this.LockMoveY);
            yield return new CellTuple(nameof(SRCCON.LockRotate), SRCCON.LockRotate, this.LockRotate);
            yield return new CellTuple(nameof(SRCCON.LockSelect), SRCCON.LockSelect, this.LockSelect);
            yield return new CellTuple(nameof(SRCCON.LockTextEdit), SRCCON.LockTextEdit, this.LockTextEdit);
            yield return new CellTuple(nameof(SRCCON.LockThemeColors), SRCCON.LockThemeColors, this.LockThemeColors);
            yield return new CellTuple(nameof(SRCCON.LockThemeEffects), SRCCON.LockThemeEffects, this.LockThemeEffects);
            yield return new CellTuple(nameof(SRCCON.LockVertexEdit), SRCCON.LockVertexEdit, this.LockVertexEdit);
            yield return new CellTuple(nameof(SRCCON.LockWidth), SRCCON.LockWidth, this.LockWidth);
            yield return new CellTuple(nameof(SRCCON.OneDBeginX), SRCCON.OneDBeginX, this.OneDBeginX);
            yield return new CellTuple(nameof(SRCCON.OneDBeginY), SRCCON.OneDBeginY, this.OneDBeginY);
            yield return new CellTuple(nameof(SRCCON.OneDEndX), SRCCON.OneDEndX, this.OneDEndX);
            yield return new CellTuple(nameof(SRCCON.OneDEndY), SRCCON.OneDEndY, this.OneDEndY);
            yield return new CellTuple(nameof(SRCCON.PageShadowObliqueAngle), SRCCON.PageShadowObliqueAngle, this.PageShadowObliqueAngle);
            yield return new CellTuple(nameof(SRCCON.PageShadowOffsetX), SRCCON.PageShadowOffsetX, this.PageShadowOffsetX);
            yield return new CellTuple(nameof(SRCCON.PageShadowOffsetY), SRCCON.PageShadowOffsetY, this.PageShadowOffsetY);
            yield return new CellTuple(nameof(SRCCON.PageShadowScaleFactor), SRCCON.PageShadowScaleFactor, this.PageShadowScaleFactor);
            yield return new CellTuple(nameof(SRCCON.PageShadowType), SRCCON.PageShadowType, this.PageShadowType);
            yield return new CellTuple(nameof(SRCCON.TextXFormAngle), SRCCON.TextXFormAngle, this.TextXFormAngle);
            yield return new CellTuple(nameof(SRCCON.TextXFormHeight), SRCCON.TextXFormHeight, this.TextXFormHeight);
            yield return new CellTuple(nameof(SRCCON.TextXFormLocPinX), SRCCON.TextXFormLocPinX, this.TextXFormLocPinX);
            yield return new CellTuple(nameof(SRCCON.TextXFormLocPinY), SRCCON.TextXFormLocPinY, this.TextXFormLocPinY);
            yield return new CellTuple(nameof(SRCCON.TextXFormPinX), SRCCON.TextXFormPinX, this.TextXFormPinX);
            yield return new CellTuple(nameof(SRCCON.TextXFormPinY), SRCCON.TextXFormPinY, this.TextXFormPinY);
            yield return new CellTuple(nameof(SRCCON.TextXFormWidth), SRCCON.TextXFormWidth, this.TextXFormWidth);
            yield return new CellTuple(nameof(SRCCON.XFormAngle), SRCCON.XFormAngle, this.XFormAngle);
            yield return new CellTuple(nameof(SRCCON.XFormHeight), SRCCON.XFormHeight, this.XFormHeight);
            yield return new CellTuple(nameof(SRCCON.XFormLocPinX), SRCCON.XFormLocPinX, this.XFormLocPinX);
            yield return new CellTuple(nameof(SRCCON.XFormLocPinY), SRCCON.XFormLocPinY, this.XFormLocPinY);
            yield return new CellTuple(nameof(SRCCON.XFormPinX), SRCCON.XFormPinX, this.XFormPinX);
            yield return new CellTuple(nameof(SRCCON.XFormPinY), SRCCON.XFormPinY, this.XFormPinY);
            yield return new CellTuple(nameof(SRCCON.XFormWidth), SRCCON.XFormWidth, this.XFormWidth);
        }
    }
}


/*

Shape Cells
     [


    'CharCase',
    'CharColor',
    'CharColorTransparency',
    'CharFont',
    'CharFontScale',
    'CharLetterspace',
    'CharSize',
    'CharStyle',
    'FillBackground',
    'FillBackgroundTransparency',
    'FillForeground',
    'FillForegroundTransparency',
    'FillPattern',
    'FillShadowBackground',
    'FillShadowBackgroundTransparency',
    'FillShadowForeground',
    'FillShadowForegroundTransparency',
    'FillShadowPattern',
    'GroupSelectMode',
    'LineBeginArrow',
    'LineBeginArrowSize',
    'LineCap',
    'LineColor',
    'LineEndArrow',
    'LineEndArrowSize',
    'LinePattern',
    'LineRounding',
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
    'OneDBeginX',
    'OneDBeginY',
    'OneDEndX',
    'OneDEndY',
    'PageShadowObliqueAngle',
    'PageShadowOffsetX',
    'PageShadowOffsetY',
    'PageShadowScaleFactor',
    'PageShadowType',
    'TextXFormAngle',
    'TextXFormHeight',
    'TextXFormLocPinX',
    'TextXFormLocPinY',
    'TextXFormPinX',
    'TextXFormPinY',
    'TextXFormWidth',
    'XFormAngle',
    'XFormHeight',
    'XFormLocPinX',
    'XFormLocPinY',
    'XFormPinX',
    'XFormPinY',
    'XFormWidth'
    ]

     */
