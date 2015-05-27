using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioPowerShell
{
    public class CellSRCDictionary : CellNameDictionary<SRC>
    {
        private static CellSRCDictionary shape_cellmap;
        private static CellSRCDictionary page_cellmap;

        public static CellSRCDictionary GetCellMapForShapes()
        {
            if (CellSRCDictionary.shape_cellmap == null)
            {
                CellSRCDictionary.shape_cellmap = new CellSRCDictionary();






                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.Angle)] = SRCConstants.Angle;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.BeginX)] = SRCConstants.BeginX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.BeginY)] = SRCConstants.BeginY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.BeginArrow)] = SRCConstants.BeginArrow;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.BeginArrowSize)] = SRCConstants.BeginArrowSize;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharCase)] = SRCConstants.CharCase;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharColor)] = SRCConstants.CharColor;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharColorTrans)] = SRCConstants.CharColorTrans;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharFont)] = SRCConstants.CharFont;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharFontScale)] = SRCConstants.CharFontScale;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharLetterspace)] = SRCConstants.CharLetterspace;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharSize)] = SRCConstants.CharSize;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.CharStyle)] = SRCConstants.CharStyle;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.EndX)] = SRCConstants.EndX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.EndY)] = SRCConstants.EndY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.EndArrow)] = SRCConstants.EndArrow;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.EndArrowSize)] = SRCConstants.EndArrowSize;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.FillBkgnd)] = SRCConstants.FillBkgnd;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.FillBkgndTrans)] = SRCConstants.FillBkgndTrans;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.FillForegnd)] = SRCConstants.FillForegnd;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.FillForegndTrans)] = SRCConstants.FillForegndTrans;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.FillPattern)] = SRCConstants.FillPattern;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.Height)] = SRCConstants.Height;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LineCap)] = SRCConstants.LineCap;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LineColor)] = SRCConstants.LineColor;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LinePattern)] = SRCConstants.LinePattern;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LineWeight)] = SRCConstants.LineWeight;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockAspect)] = SRCConstants.LockAspect;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockBegin)] = SRCConstants.LockBegin;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockCalcWH)] = SRCConstants.LockCalcWH;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockCrop)] = SRCConstants.LockCrop;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockCustProp)] = SRCConstants.LockCustProp;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockDelete)] = SRCConstants.LockDelete;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockEnd)] = SRCConstants.LockEnd;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockFormat)] = SRCConstants.LockFormat;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockFromGroupFormat)] = SRCConstants.LockFromGroupFormat;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockGroup)] = SRCConstants.LockGroup;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockHeight)] = SRCConstants.LockHeight;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockMoveX)] = SRCConstants.LockMoveX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockMoveY)] = SRCConstants.LockMoveY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockRotate)] = SRCConstants.LockRotate;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockSelect)] = SRCConstants.LockSelect;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockTextEdit)] = SRCConstants.LockTextEdit;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockThemeColors)] = SRCConstants.LockThemeColors;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockThemeEffects)] = SRCConstants.LockThemeEffects;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockVtxEdit)] = SRCConstants.LockVtxEdit;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LockWidth)] = SRCConstants.LockWidth;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LocPinX)] = SRCConstants.LocPinX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.LocPinY)] = SRCConstants.LocPinY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.PinX)] = SRCConstants.PinX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.PinY)] = SRCConstants.PinY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.Rounding)] = SRCConstants.Rounding;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.SelectMode)] = SRCConstants.SelectMode;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwBkgnd)] = SRCConstants.ShdwBkgnd;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwBkgndTrans)] = SRCConstants.ShdwBkgndTrans;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwForegnd)] = SRCConstants.ShdwForegnd;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwForegndTrans)] = SRCConstants.ShdwForegndTrans;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwObliqueAngle)] = SRCConstants.ShdwObliqueAngle;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwOffsetX)] = SRCConstants.ShdwOffsetX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwOffsetY)] = SRCConstants.ShdwOffsetY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwPattern)] = SRCConstants.ShdwPattern;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwScaleFactor)] = SRCConstants.ShdwScaleFactor;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.ShdwType)] = SRCConstants.ShdwType;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtAngle)] = SRCConstants.TxtAngle;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtHeight)] = SRCConstants.TxtHeight;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtLocPinX)] = SRCConstants.TxtLocPinX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtLocPinY)] = SRCConstants.TxtLocPinY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtPinX)] = SRCConstants.TxtPinX;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtPinY)] = SRCConstants.TxtPinY;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.TxtWidth)] = SRCConstants.TxtWidth;
                CellSRCDictionary.shape_cellmap[nameof(SRCConstants.Width)] = SRCConstants.Width;

            }
            return CellSRCDictionary.shape_cellmap;
        }

        public static CellSRCDictionary GetCellMapForPages()
        {
            if (CellSRCDictionary.page_cellmap == null)
            {
                CellSRCDictionary.page_cellmap = new CellSRCDictionary();
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.PageBottomMargin)] = SRCConstants.PageBottomMargin;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.PageHeight)] = SRCConstants.PageHeight;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.PageLeftMargin)] = SRCConstants.PageLeftMargin;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.PageLineJumpDirX)] = SRCConstants.PageLineJumpDirX;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.PageLineJumpDirY)] = SRCConstants.PageLineJumpDirY;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.PageRightMargin)] = SRCConstants.PageRightMargin;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.PageScale)] = SRCConstants.PageScale;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.PageShapeSplit)] = SRCConstants.PageShapeSplit;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.PageTopMargin)] = SRCConstants.PageTopMargin;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.PageWidth)] = SRCConstants.PageWidth;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.CenterX)] = SRCConstants.CenterX;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.CenterY)] = SRCConstants.CenterY;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.PaperKind)] = SRCConstants.PaperKind;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.PrintGrid)] = SRCConstants.PrintGrid;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.PrintPageOrientation)] = SRCConstants.PrintPageOrientation;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.ScaleX)] = SRCConstants.ScaleX;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.ScaleY)] = SRCConstants.ScaleY;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.PaperSource)] = SRCConstants.PaperSource;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.DrawingScale)] = SRCConstants.DrawingScale;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.DrawingScaleType)] = SRCConstants.DrawingScaleType;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.DrawingSizeType)] = SRCConstants.DrawingSizeType;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.InhibitSnap)] = SRCConstants.InhibitSnap;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.ShdwObliqueAngle)] = SRCConstants.ShdwObliqueAngle;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.ShdwOffsetX)] = SRCConstants.ShdwOffsetX;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.ShdwOffsetY)] = SRCConstants.ShdwOffsetY;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.ShdwScaleFactor)] = SRCConstants.ShdwScaleFactor;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.ShdwType)] = SRCConstants.ShdwType;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.UIVisibility)] = SRCConstants.UIVisibility;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.XGridDensity)] = SRCConstants.XGridDensity;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.XGridOrigin)] = SRCConstants.XGridOrigin;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.XGridSpacing)] = SRCConstants.XGridSpacing;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.XRulerDensity)] = SRCConstants.XRulerDensity;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.XRulerOrigin)] = SRCConstants.XRulerOrigin;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.YGridDensity)] = SRCConstants.YGridDensity;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.YGridOrigin)] = SRCConstants.YGridOrigin;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.YGridSpacing)] = SRCConstants.YGridSpacing;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.YRulerDensity)] = SRCConstants.YRulerDensity;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.YRulerOrigin)] = SRCConstants.YRulerOrigin;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.AvenueSizeX)] = SRCConstants.AvenueSizeX;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.AvenueSizeY)] = SRCConstants.AvenueSizeY;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.BlockSizeX)] = SRCConstants.BlockSizeX;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.BlockSizeY)] = SRCConstants.BlockSizeY;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.CtrlAsInput)] = SRCConstants.CtrlAsInput;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.DynamicsOff)] = SRCConstants.DynamicsOff;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.EnableGrid)] = SRCConstants.EnableGrid;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.LineAdjustFrom)] = SRCConstants.LineAdjustFrom;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.LineAdjustTo)] = SRCConstants.LineAdjustTo;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.LineJumpCode)] = SRCConstants.LineJumpCode;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.LineJumpFactorX)] = SRCConstants.LineJumpFactorX;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.LineJumpFactorY)] = SRCConstants.LineJumpFactorY;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.LineJumpStyle)] = SRCConstants.LineJumpStyle;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.LineRouteExt)] = SRCConstants.LineRouteExt;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.LineToLineX)] = SRCConstants.LineToLineX;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.LineToLineY)] = SRCConstants.LineToLineY;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.LineToNodeX)] = SRCConstants.LineToNodeX;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.LineToNodeY)] = SRCConstants.LineToNodeY;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.PlaceDepth)] = SRCConstants.PlaceDepth;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.PlaceFlip)] = SRCConstants.PlaceFlip;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.PlaceStyle)] = SRCConstants.PlaceStyle;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.PlowCode)] = SRCConstants.PlowCode;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.ResizePage)] = SRCConstants.ResizePage;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.RouteStyle)] = SRCConstants.RouteStyle;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.AvoidPageBreaks)] = SRCConstants.AvoidPageBreaks;
                CellSRCDictionary.page_cellmap[nameof(SRCConstants.DrawingResizeType)] = SRCConstants.DrawingResizeType;
            }
            return CellSRCDictionary.page_cellmap;
        }

        public CellQuery CreateQueryFromCellNames(IList<string> Cells)
        {
            var invalid_names = Cells.Where(cellname => !this.ContainsCell(cellname)).ToList();
            if (invalid_names.Count > 0)
            {
                string msg = "Invalid cell names: " + string.Join(",", invalid_names);
                throw new ArgumentException(msg);
            }

            var query = new CellQuery();

            foreach (string resolved_cellname in this.ResolveNames(Cells))
            {
                if (!query.CellColumns.Contains(resolved_cellname))
                {
                    var resolved_src = this[resolved_cellname];
                    query.AddCell(resolved_src, resolved_cellname);
                }
            }
            return query;
        }
    }
}