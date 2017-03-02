using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Scripting.ShapeSheet
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
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.Angle)] = SrcConstants.Angle;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.BeginX)] = SrcConstants.BeginX;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.BeginY)] = SrcConstants.BeginY;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.BeginArrow)] = SrcConstants.BeginArrow;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.BeginArrowSize)] = SrcConstants.BeginArrowSize;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.CharCase)] = SrcConstants.CharCase;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.CharColor)] = SrcConstants.CharColor;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.CharColorTrans)] = SrcConstants.CharColorTrans;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.CharFont)] = SrcConstants.CharFont;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.CharFontScale)] = SrcConstants.CharFontScale;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.CharLetterspace)] = SrcConstants.CharLetterspace;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.CharSize)] = SrcConstants.CharSize;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.CharStyle)] = SrcConstants.CharStyle;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.EndX)] = SrcConstants.EndX;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.EndY)] = SrcConstants.EndY;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.EndArrow)] = SrcConstants.EndArrow;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.EndArrowSize)] = SrcConstants.EndArrowSize;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.FillBkgnd)] = SrcConstants.FillBkgnd;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.FillBkgndTrans)] = SrcConstants.FillBkgndTrans;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.FillForegnd)] = SrcConstants.FillForegnd;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.FillForegndTrans)] = SrcConstants.FillForegndTrans;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.FillPattern)] = SrcConstants.FillPattern;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.Height)] = SrcConstants.Height;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LineCap)] = SrcConstants.LineCap;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LineColor)] = SrcConstants.LineColor;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LinePattern)] = SrcConstants.LinePattern;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LineWeight)] = SrcConstants.LineWeight;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockAspect)] = SrcConstants.LockAspect;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockBegin)] = SrcConstants.LockBegin;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockCalcWH)] = SrcConstants.LockCalcWH;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockCrop)] = SrcConstants.LockCrop;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockCustProp)] = SrcConstants.LockCustProp;
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
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockVtxEdit)] = SrcConstants.LockVtxEdit;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LockWidth)] = SrcConstants.LockWidth;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LocPinX)] = SrcConstants.LocPinX;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.LocPinY)] = SrcConstants.LocPinY;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.PinX)] = SrcConstants.PinX;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.PinY)] = SrcConstants.PinY;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.Rounding)] = SrcConstants.Rounding;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.SelectMode)] = SrcConstants.SelectMode;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.ShdwBkgnd)] = SrcConstants.ShdwBkgnd;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.ShdwBkgndTrans)] = SrcConstants.ShdwBkgndTrans;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.ShdwForegnd)] = SrcConstants.ShdwForegnd;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.ShdwForegndTrans)] = SrcConstants.ShdwForegndTrans;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.ShdwObliqueAngle)] = SrcConstants.ShdwObliqueAngle;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.ShdwOffsetX)] = SrcConstants.ShdwOffsetX;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.ShdwOffsetY)] = SrcConstants.ShdwOffsetY;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.ShdwPattern)] = SrcConstants.ShdwPattern;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.ShdwScaleFactor)] = SrcConstants.ShdwScaleFactor;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.ShdwType)] = SrcConstants.ShdwType;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.TxtAngle)] = SrcConstants.TxtAngle;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.TxtHeight)] = SrcConstants.TxtHeight;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.TxtLocPinX)] = SrcConstants.TxtLocPinX;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.TxtLocPinY)] = SrcConstants.TxtLocPinY;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.TxtPinX)] = SrcConstants.TxtPinX;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.TxtPinY)] = SrcConstants.TxtPinY;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.TxtWidth)] = SrcConstants.TxtWidth;
                CellSrcDictionary.shape_cellmap[nameof(SrcConstants.Width)] = SrcConstants.Width;

            }
            return CellSrcDictionary.shape_cellmap;
        }

        public static CellSrcDictionary GetCellMapForPages()
        {
            if (CellSrcDictionary.page_cellmap == null)
            {
                CellSrcDictionary.page_cellmap = new CellSrcDictionary();
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageBottomMargin)] = SrcConstants.PageBottomMargin;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageHeight)] = SrcConstants.PageHeight;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLeftMargin)] = SrcConstants.PageLeftMargin;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLineJumpDirX)] = SrcConstants.PageLineJumpDirX;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageLineJumpDirY)] = SrcConstants.PageLineJumpDirY;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageRightMargin)] = SrcConstants.PageRightMargin;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageScale)] = SrcConstants.PageScale;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageShapeSplit)] = SrcConstants.PageShapeSplit;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageTopMargin)] = SrcConstants.PageTopMargin;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PageWidth)] = SrcConstants.PageWidth;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.CenterX)] = SrcConstants.CenterX;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.CenterY)] = SrcConstants.CenterY;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PaperKind)] = SrcConstants.PaperKind;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PrintGrid)] = SrcConstants.PrintGrid;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PrintPageOrientation)] = SrcConstants.PrintPageOrientation;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.ScaleX)] = SrcConstants.ScaleX;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.ScaleY)] = SrcConstants.ScaleY;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PaperSource)] = SrcConstants.PaperSource;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.DrawingScale)] = SrcConstants.DrawingScale;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.DrawingScaleType)] = SrcConstants.DrawingScaleType;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.DrawingSizeType)] = SrcConstants.DrawingSizeType;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.InhibitSnap)] = SrcConstants.InhibitSnap;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.ShdwObliqueAngle)] = SrcConstants.ShdwObliqueAngle;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.ShdwOffsetX)] = SrcConstants.ShdwOffsetX;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.ShdwOffsetY)] = SrcConstants.ShdwOffsetY;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.ShdwScaleFactor)] = SrcConstants.ShdwScaleFactor;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.ShdwType)] = SrcConstants.ShdwType;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.UIVisibility)] = SrcConstants.UIVisibility;
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
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.AvenueSizeX)] = SrcConstants.AvenueSizeX;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.AvenueSizeY)] = SrcConstants.AvenueSizeY;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.BlockSizeX)] = SrcConstants.BlockSizeX;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.BlockSizeY)] = SrcConstants.BlockSizeY;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.CtrlAsInput)] = SrcConstants.CtrlAsInput;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.DynamicsOff)] = SrcConstants.DynamicsOff;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.EnableGrid)] = SrcConstants.EnableGrid;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.LineAdjustFrom)] = SrcConstants.LineAdjustFrom;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.LineAdjustTo)] = SrcConstants.LineAdjustTo;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.LineJumpCode)] = SrcConstants.LineJumpCode;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.LineJumpFactorX)] = SrcConstants.LineJumpFactorX;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.LineJumpFactorY)] = SrcConstants.LineJumpFactorY;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.LineJumpStyle)] = SrcConstants.LineJumpStyle;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.LineRouteExt)] = SrcConstants.LineRouteExt;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.LineToLineX)] = SrcConstants.LineToLineX;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.LineToLineY)] = SrcConstants.LineToLineY;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.LineToNodeX)] = SrcConstants.LineToNodeX;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.LineToNodeY)] = SrcConstants.LineToNodeY;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PlaceDepth)] = SrcConstants.PlaceDepth;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PlaceFlip)] = SrcConstants.PlaceFlip;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PlaceStyle)] = SrcConstants.PlaceStyle;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.PlowCode)] = SrcConstants.PlowCode;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.ResizePage)] = SrcConstants.ResizePage;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.RouteStyle)] = SrcConstants.RouteStyle;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.AvoidPageBreaks)] = SrcConstants.AvoidPageBreaks;
                CellSrcDictionary.page_cellmap[nameof(SrcConstants.DrawingResizeType)] = SrcConstants.DrawingResizeType;
            }
            return CellSrcDictionary.page_cellmap;
        }
    }
}

