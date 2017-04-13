using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Scripting.Models
{
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

