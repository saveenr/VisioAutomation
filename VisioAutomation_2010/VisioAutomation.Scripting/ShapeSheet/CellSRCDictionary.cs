using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Scripting.ShapeSheet
{
    public class CellSRCDictionary : NamedSRCDictionary
    {
        private static CellSRCDictionary shape_cellmap;
        private static CellSRCDictionary page_cellmap;

        public static CellSRCDictionary GetCellMapForShapes()
        {
            if (CellSRCDictionary.shape_cellmap == null)
            {
                CellSRCDictionary.shape_cellmap = new CellSRCDictionary();
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.Angle)] = SrcConstants.Angle;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.BeginX)] = SrcConstants.BeginX;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.BeginY)] = SrcConstants.BeginY;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.BeginArrow)] = SrcConstants.BeginArrow;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.BeginArrowSize)] = SrcConstants.BeginArrowSize;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.CharCase)] = SrcConstants.CharCase;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.CharColor)] = SrcConstants.CharColor;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.CharColorTrans)] = SrcConstants.CharColorTrans;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.CharFont)] = SrcConstants.CharFont;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.CharFontScale)] = SrcConstants.CharFontScale;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.CharLetterspace)] = SrcConstants.CharLetterspace;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.CharSize)] = SrcConstants.CharSize;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.CharStyle)] = SrcConstants.CharStyle;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.EndX)] = SrcConstants.EndX;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.EndY)] = SrcConstants.EndY;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.EndArrow)] = SrcConstants.EndArrow;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.EndArrowSize)] = SrcConstants.EndArrowSize;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.FillBkgnd)] = SrcConstants.FillBkgnd;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.FillBkgndTrans)] = SrcConstants.FillBkgndTrans;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.FillForegnd)] = SrcConstants.FillForegnd;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.FillForegndTrans)] = SrcConstants.FillForegndTrans;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.FillPattern)] = SrcConstants.FillPattern;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.Height)] = SrcConstants.Height;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LineCap)] = SrcConstants.LineCap;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LineColor)] = SrcConstants.LineColor;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LinePattern)] = SrcConstants.LinePattern;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LineWeight)] = SrcConstants.LineWeight;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockAspect)] = SrcConstants.LockAspect;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockBegin)] = SrcConstants.LockBegin;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockCalcWH)] = SrcConstants.LockCalcWH;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockCrop)] = SrcConstants.LockCrop;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockCustProp)] = SrcConstants.LockCustProp;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockDelete)] = SrcConstants.LockDelete;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockEnd)] = SrcConstants.LockEnd;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockFormat)] = SrcConstants.LockFormat;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockFromGroupFormat)] = SrcConstants.LockFromGroupFormat;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockGroup)] = SrcConstants.LockGroup;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockHeight)] = SrcConstants.LockHeight;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockMoveX)] = SrcConstants.LockMoveX;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockMoveY)] = SrcConstants.LockMoveY;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockRotate)] = SrcConstants.LockRotate;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockSelect)] = SrcConstants.LockSelect;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockTextEdit)] = SrcConstants.LockTextEdit;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockThemeColors)] = SrcConstants.LockThemeColors;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockThemeEffects)] = SrcConstants.LockThemeEffects;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockVtxEdit)] = SrcConstants.LockVtxEdit;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LockWidth)] = SrcConstants.LockWidth;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LocPinX)] = SrcConstants.LocPinX;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.LocPinY)] = SrcConstants.LocPinY;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.PinX)] = SrcConstants.PinX;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.PinY)] = SrcConstants.PinY;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.Rounding)] = SrcConstants.Rounding;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.SelectMode)] = SrcConstants.SelectMode;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.ShdwBkgnd)] = SrcConstants.ShdwBkgnd;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.ShdwBkgndTrans)] = SrcConstants.ShdwBkgndTrans;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.ShdwForegnd)] = SrcConstants.ShdwForegnd;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.ShdwForegndTrans)] = SrcConstants.ShdwForegndTrans;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.ShdwObliqueAngle)] = SrcConstants.ShdwObliqueAngle;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.ShdwOffsetX)] = SrcConstants.ShdwOffsetX;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.ShdwOffsetY)] = SrcConstants.ShdwOffsetY;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.ShdwPattern)] = SrcConstants.ShdwPattern;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.ShdwScaleFactor)] = SrcConstants.ShdwScaleFactor;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.ShdwType)] = SrcConstants.ShdwType;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.TxtAngle)] = SrcConstants.TxtAngle;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.TxtHeight)] = SrcConstants.TxtHeight;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.TxtLocPinX)] = SrcConstants.TxtLocPinX;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.TxtLocPinY)] = SrcConstants.TxtLocPinY;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.TxtPinX)] = SrcConstants.TxtPinX;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.TxtPinY)] = SrcConstants.TxtPinY;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.TxtWidth)] = SrcConstants.TxtWidth;
                CellSRCDictionary.shape_cellmap[nameof(SrcConstants.Width)] = SrcConstants.Width;

            }
            return CellSRCDictionary.shape_cellmap;
        }

        public static CellSRCDictionary GetCellMapForPages()
        {
            if (CellSRCDictionary.page_cellmap == null)
            {
                CellSRCDictionary.page_cellmap = new CellSRCDictionary();
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.PageBottomMargin)] = SrcConstants.PageBottomMargin;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.PageHeight)] = SrcConstants.PageHeight;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.PageLeftMargin)] = SrcConstants.PageLeftMargin;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.PageLineJumpDirX)] = SrcConstants.PageLineJumpDirX;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.PageLineJumpDirY)] = SrcConstants.PageLineJumpDirY;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.PageRightMargin)] = SrcConstants.PageRightMargin;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.PageScale)] = SrcConstants.PageScale;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.PageShapeSplit)] = SrcConstants.PageShapeSplit;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.PageTopMargin)] = SrcConstants.PageTopMargin;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.PageWidth)] = SrcConstants.PageWidth;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.CenterX)] = SrcConstants.CenterX;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.CenterY)] = SrcConstants.CenterY;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.PaperKind)] = SrcConstants.PaperKind;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.PrintGrid)] = SrcConstants.PrintGrid;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.PrintPageOrientation)] = SrcConstants.PrintPageOrientation;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.ScaleX)] = SrcConstants.ScaleX;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.ScaleY)] = SrcConstants.ScaleY;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.PaperSource)] = SrcConstants.PaperSource;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.DrawingScale)] = SrcConstants.DrawingScale;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.DrawingScaleType)] = SrcConstants.DrawingScaleType;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.DrawingSizeType)] = SrcConstants.DrawingSizeType;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.InhibitSnap)] = SrcConstants.InhibitSnap;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.ShdwObliqueAngle)] = SrcConstants.ShdwObliqueAngle;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.ShdwOffsetX)] = SrcConstants.ShdwOffsetX;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.ShdwOffsetY)] = SrcConstants.ShdwOffsetY;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.ShdwScaleFactor)] = SrcConstants.ShdwScaleFactor;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.ShdwType)] = SrcConstants.ShdwType;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.UIVisibility)] = SrcConstants.UIVisibility;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.XGridDensity)] = SrcConstants.XGridDensity;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.XGridOrigin)] = SrcConstants.XGridOrigin;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.XGridSpacing)] = SrcConstants.XGridSpacing;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.XRulerDensity)] = SrcConstants.XRulerDensity;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.XRulerOrigin)] = SrcConstants.XRulerOrigin;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.YGridDensity)] = SrcConstants.YGridDensity;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.YGridOrigin)] = SrcConstants.YGridOrigin;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.YGridSpacing)] = SrcConstants.YGridSpacing;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.YRulerDensity)] = SrcConstants.YRulerDensity;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.YRulerOrigin)] = SrcConstants.YRulerOrigin;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.AvenueSizeX)] = SrcConstants.AvenueSizeX;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.AvenueSizeY)] = SrcConstants.AvenueSizeY;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.BlockSizeX)] = SrcConstants.BlockSizeX;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.BlockSizeY)] = SrcConstants.BlockSizeY;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.CtrlAsInput)] = SrcConstants.CtrlAsInput;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.DynamicsOff)] = SrcConstants.DynamicsOff;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.EnableGrid)] = SrcConstants.EnableGrid;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.LineAdjustFrom)] = SrcConstants.LineAdjustFrom;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.LineAdjustTo)] = SrcConstants.LineAdjustTo;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.LineJumpCode)] = SrcConstants.LineJumpCode;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.LineJumpFactorX)] = SrcConstants.LineJumpFactorX;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.LineJumpFactorY)] = SrcConstants.LineJumpFactorY;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.LineJumpStyle)] = SrcConstants.LineJumpStyle;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.LineRouteExt)] = SrcConstants.LineRouteExt;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.LineToLineX)] = SrcConstants.LineToLineX;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.LineToLineY)] = SrcConstants.LineToLineY;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.LineToNodeX)] = SrcConstants.LineToNodeX;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.LineToNodeY)] = SrcConstants.LineToNodeY;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.PlaceDepth)] = SrcConstants.PlaceDepth;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.PlaceFlip)] = SrcConstants.PlaceFlip;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.PlaceStyle)] = SrcConstants.PlaceStyle;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.PlowCode)] = SrcConstants.PlowCode;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.ResizePage)] = SrcConstants.ResizePage;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.RouteStyle)] = SrcConstants.RouteStyle;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.AvoidPageBreaks)] = SrcConstants.AvoidPageBreaks;
                CellSRCDictionary.page_cellmap[nameof(SrcConstants.DrawingResizeType)] = SrcConstants.DrawingResizeType;
            }
            return CellSRCDictionary.page_cellmap;
        }
    }
}

