using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using VA = VisioAutomation;

namespace VisioPowerShell
{
    public class CellSRCDictionary : CellNameDictionary<SRC>
    {
        private static CellSRCDictionary shape_cellmap;
        private static CellSRCDictionary page_cellmap;

        public CellSRCDictionary() :
            base()
        {
        }

        public static CellSRCDictionary GetCellMapForShapes()
        {
            if (CellSRCDictionary.shape_cellmap == null)
            {
                CellSRCDictionary.shape_cellmap = new CellSRCDictionary();
                CellSRCDictionary.shape_cellmap["Angle"] = SRCConstants.Angle;
                CellSRCDictionary.shape_cellmap["BeginX"] = SRCConstants.BeginX;
                CellSRCDictionary.shape_cellmap["BeginY"] = SRCConstants.BeginY;
                CellSRCDictionary.shape_cellmap["BeginArrow"] = SRCConstants.BeginArrow;
                CellSRCDictionary.shape_cellmap["BeginArrowSize"] = SRCConstants.BeginArrowSize;
                CellSRCDictionary.shape_cellmap["CharCase"] = SRCConstants.CharCase;
                CellSRCDictionary.shape_cellmap["CharColor"] = SRCConstants.CharColor;
                CellSRCDictionary.shape_cellmap["CharColorTrans"] = SRCConstants.CharColorTrans;
                CellSRCDictionary.shape_cellmap["CharFont"] = SRCConstants.CharFont;
                CellSRCDictionary.shape_cellmap["CharFontScale"] = SRCConstants.CharFontScale;
                CellSRCDictionary.shape_cellmap["CharLetterspace"] = SRCConstants.CharLetterspace;
                CellSRCDictionary.shape_cellmap["CharSize"] = SRCConstants.CharSize;
                CellSRCDictionary.shape_cellmap["CharStyle"] = SRCConstants.CharStyle;
                CellSRCDictionary.shape_cellmap["EndX"] = SRCConstants.EndX;
                CellSRCDictionary.shape_cellmap["EndY"] = SRCConstants.EndY;
                CellSRCDictionary.shape_cellmap["EndArrow"] = SRCConstants.EndArrow;
                CellSRCDictionary.shape_cellmap["EndArrowSize"] = SRCConstants.EndArrowSize;
                CellSRCDictionary.shape_cellmap["FillBkgnd"] = SRCConstants.FillBkgnd;
                CellSRCDictionary.shape_cellmap["FillBkgndTrans"] = SRCConstants.FillBkgndTrans;
                CellSRCDictionary.shape_cellmap["FillForegnd"] = SRCConstants.FillForegnd;
                CellSRCDictionary.shape_cellmap["FillForegndTrans"] = SRCConstants.FillForegndTrans;
                CellSRCDictionary.shape_cellmap["FillPattern"] = SRCConstants.FillPattern;
                CellSRCDictionary.shape_cellmap["Height"] = SRCConstants.Height;
                CellSRCDictionary.shape_cellmap["LineCap"] = SRCConstants.LineCap;
                CellSRCDictionary.shape_cellmap["LineColor"] = SRCConstants.LineColor;
                CellSRCDictionary.shape_cellmap["LinePattern"] = SRCConstants.LinePattern;
                CellSRCDictionary.shape_cellmap["LineWeight"] = SRCConstants.LineWeight;
                CellSRCDictionary.shape_cellmap["LockAspect"] = SRCConstants.LockAspect;
                CellSRCDictionary.shape_cellmap["LockBegin"] = SRCConstants.LockBegin;
                CellSRCDictionary.shape_cellmap["LockCalcWH"] = SRCConstants.LockCalcWH;
                CellSRCDictionary.shape_cellmap["LockCrop"] = SRCConstants.LockCrop;
                CellSRCDictionary.shape_cellmap["LockCustProp"] = SRCConstants.LockCustProp;
                CellSRCDictionary.shape_cellmap["LockDelete"] = SRCConstants.LockDelete;
                CellSRCDictionary.shape_cellmap["LockEnd"] = SRCConstants.LockEnd;
                CellSRCDictionary.shape_cellmap["LockFormat"] = SRCConstants.LockFormat;
                CellSRCDictionary.shape_cellmap["LockFromGroupFormat"] = SRCConstants.LockFromGroupFormat;
                CellSRCDictionary.shape_cellmap["LockGroup"] = SRCConstants.LockGroup;
                CellSRCDictionary.shape_cellmap["LockHeight"] = SRCConstants.LockHeight;
                CellSRCDictionary.shape_cellmap["LockMoveX"] = SRCConstants.LockMoveX;
                CellSRCDictionary.shape_cellmap["LockMoveY"] = SRCConstants.LockMoveY;
                CellSRCDictionary.shape_cellmap["LockRotate"] = SRCConstants.LockRotate;
                CellSRCDictionary.shape_cellmap["LockSelect"] = SRCConstants.LockSelect;
                CellSRCDictionary.shape_cellmap["LockTextEdit"] = SRCConstants.LockTextEdit;
                CellSRCDictionary.shape_cellmap["LockThemeColors"] = SRCConstants.LockThemeColors;
                CellSRCDictionary.shape_cellmap["LockThemeEffects"] = SRCConstants.LockThemeEffects;
                CellSRCDictionary.shape_cellmap["LockVtxEdit"] = SRCConstants.LockVtxEdit;
                CellSRCDictionary.shape_cellmap["LockWidth"] = SRCConstants.LockWidth;
                CellSRCDictionary.shape_cellmap["LocPinX"] = SRCConstants.LocPinX;
                CellSRCDictionary.shape_cellmap["LocPinY"] = SRCConstants.LocPinY;
                CellSRCDictionary.shape_cellmap["PinX"] = SRCConstants.PinX;
                CellSRCDictionary.shape_cellmap["PinY"] = SRCConstants.PinY;
                CellSRCDictionary.shape_cellmap["Rounding"] = SRCConstants.Rounding;
                CellSRCDictionary.shape_cellmap["SelectMode"] = SRCConstants.SelectMode;
                CellSRCDictionary.shape_cellmap["ShdwBkgnd"] = SRCConstants.ShdwBkgnd;
                CellSRCDictionary.shape_cellmap["ShdwBkgndTrans"] = SRCConstants.ShdwBkgndTrans;
                CellSRCDictionary.shape_cellmap["ShdwForegnd"] = SRCConstants.ShdwForegnd;
                CellSRCDictionary.shape_cellmap["ShdwForegndTrans"] = SRCConstants.ShdwForegndTrans;
                CellSRCDictionary.shape_cellmap["ShdwObliqueAngle"] = SRCConstants.ShdwObliqueAngle;
                CellSRCDictionary.shape_cellmap["ShdwOffsetX"] = SRCConstants.ShdwOffsetX;
                CellSRCDictionary.shape_cellmap["ShdwOffsetY"] = SRCConstants.ShdwOffsetY;
                CellSRCDictionary.shape_cellmap["ShdwPattern"] = SRCConstants.ShdwPattern;
                CellSRCDictionary.shape_cellmap["ShdwScaleFactor"] = SRCConstants.ShdwScaleFactor;
                CellSRCDictionary.shape_cellmap["ShdwType"] = SRCConstants.ShdwType;
                CellSRCDictionary.shape_cellmap["TxtAngle"] = SRCConstants.TxtAngle;
                CellSRCDictionary.shape_cellmap["TxtHeight"] = SRCConstants.TxtHeight;
                CellSRCDictionary.shape_cellmap["TxtLocPinX"] = SRCConstants.TxtLocPinX;
                CellSRCDictionary.shape_cellmap["TxtLocPinY"] = SRCConstants.TxtLocPinY;
                CellSRCDictionary.shape_cellmap["TxtPinX"] = SRCConstants.TxtPinX;
                CellSRCDictionary.shape_cellmap["TxtPinY"] = SRCConstants.TxtPinY;
                CellSRCDictionary.shape_cellmap["TxtWidth"] = SRCConstants.TxtWidth;
                CellSRCDictionary.shape_cellmap["Width"] = SRCConstants.Width;

            }
            return CellSRCDictionary.shape_cellmap;
        }

        public static CellSRCDictionary GetCellMapForPages()
        {
            if (CellSRCDictionary.page_cellmap == null)
            {
                CellSRCDictionary.page_cellmap = new CellSRCDictionary();
                CellSRCDictionary.page_cellmap["PageBottomMargin"] = SRCConstants.PageBottomMargin;
                CellSRCDictionary.page_cellmap["PageHeight"] = SRCConstants.PageHeight;
                CellSRCDictionary.page_cellmap["PageLeftMargin"] = SRCConstants.PageLeftMargin;
                CellSRCDictionary.page_cellmap["PageLineJumpDirX"] = SRCConstants.PageLineJumpDirX;
                CellSRCDictionary.page_cellmap["PageLineJumpDirY"] = SRCConstants.PageLineJumpDirY;

                CellSRCDictionary.page_cellmap["PageRightMargin"] = SRCConstants.PageRightMargin;
                CellSRCDictionary.page_cellmap["PageScale"] = SRCConstants.PageScale;
                CellSRCDictionary.page_cellmap["PageShapeSplit"] = SRCConstants.PageShapeSplit;
                CellSRCDictionary.page_cellmap["PageTopMargin"] = SRCConstants.PageTopMargin;
                CellSRCDictionary.page_cellmap["PageWidth"] = SRCConstants.PageWidth;
                CellSRCDictionary.page_cellmap["CenterX"] = SRCConstants.CenterX;
                CellSRCDictionary.page_cellmap["CenterY"] = SRCConstants.CenterY;
                CellSRCDictionary.page_cellmap["PaperKind"] = SRCConstants.PaperKind;
                CellSRCDictionary.page_cellmap["PrintGrid"] = SRCConstants.PrintGrid;
                CellSRCDictionary.page_cellmap["PrintPageOrientation"] = SRCConstants.PrintPageOrientation;
                CellSRCDictionary.page_cellmap["ScaleX"] = SRCConstants.ScaleX;
                CellSRCDictionary.page_cellmap["ScaleY"] = SRCConstants.ScaleY;
                CellSRCDictionary.page_cellmap["PaperSource"] = SRCConstants.PaperSource;
                CellSRCDictionary.page_cellmap["DrawingScale"] = SRCConstants.DrawingScale;
                CellSRCDictionary.page_cellmap["DrawingScaleType"] = SRCConstants.DrawingScaleType;
                CellSRCDictionary.page_cellmap["DrawingSizeType"] = SRCConstants.DrawingSizeType;
                CellSRCDictionary.page_cellmap["InhibitSnap"] = SRCConstants.InhibitSnap;
                CellSRCDictionary.page_cellmap["ShdwObliqueAngle"] = SRCConstants.ShdwObliqueAngle;
                CellSRCDictionary.page_cellmap["ShdwOffsetX"] = SRCConstants.ShdwOffsetX;
                CellSRCDictionary.page_cellmap["ShdwOffsetY"] = SRCConstants.ShdwOffsetY;
                CellSRCDictionary.page_cellmap["ShdwScaleFactor"] = SRCConstants.ShdwScaleFactor;
                CellSRCDictionary.page_cellmap["ShdwType"] = SRCConstants.ShdwType;
                CellSRCDictionary.page_cellmap["UIVisibility"] = SRCConstants.UIVisibility;
                CellSRCDictionary.page_cellmap["XGridDensity"] = SRCConstants.XGridDensity;
                CellSRCDictionary.page_cellmap["XGridOrigin"] = SRCConstants.XGridOrigin;
                CellSRCDictionary.page_cellmap["XGridSpacing"] = SRCConstants.XGridSpacing;
                CellSRCDictionary.page_cellmap["XRulerDensity"] = SRCConstants.XRulerDensity;
                CellSRCDictionary.page_cellmap["XRulerOrigin"] = SRCConstants.XRulerOrigin;
                CellSRCDictionary.page_cellmap["YGridDensity"] = SRCConstants.YGridDensity;
                CellSRCDictionary.page_cellmap["YGridOrigin"] = SRCConstants.YGridOrigin;
                CellSRCDictionary.page_cellmap["YGridSpacing"] = SRCConstants.YGridSpacing;
                CellSRCDictionary.page_cellmap["YRulerDensity"] = SRCConstants.YRulerDensity;
                CellSRCDictionary.page_cellmap["YRulerOrigin"] = SRCConstants.YRulerOrigin;
                CellSRCDictionary.page_cellmap["AvenueSizeX"] = SRCConstants.AvenueSizeX;
                CellSRCDictionary.page_cellmap["AvenueSizeY"] = SRCConstants.AvenueSizeY;
                CellSRCDictionary.page_cellmap["BlockSizeX"] = SRCConstants.BlockSizeX;
                CellSRCDictionary.page_cellmap["BlockSizeY"] = SRCConstants.BlockSizeY;
                CellSRCDictionary.page_cellmap["CtrlAsInput"] = SRCConstants.CtrlAsInput;
                CellSRCDictionary.page_cellmap["DynamicsOff"] = SRCConstants.DynamicsOff;
                CellSRCDictionary.page_cellmap["EnableGrid"] = SRCConstants.EnableGrid;
                CellSRCDictionary.page_cellmap["LineAdjustFrom"] = SRCConstants.LineAdjustFrom;
                CellSRCDictionary.page_cellmap["LineAdjustTo"] = SRCConstants.LineAdjustTo;
                CellSRCDictionary.page_cellmap["LineJumpCode"] = SRCConstants.LineJumpCode;
                CellSRCDictionary.page_cellmap["LineJumpFactorX"] = SRCConstants.LineJumpFactorX;
                CellSRCDictionary.page_cellmap["LineJumpFactorY"] = SRCConstants.LineJumpFactorY;
                CellSRCDictionary.page_cellmap["LineJumpStyle"] = SRCConstants.LineJumpStyle;
                CellSRCDictionary.page_cellmap["LineRouteExt"] = SRCConstants.LineRouteExt;
                CellSRCDictionary.page_cellmap["LineToLineX"] = SRCConstants.LineToLineX;
                CellSRCDictionary.page_cellmap["LineToLineY"] = SRCConstants.LineToLineY;
                CellSRCDictionary.page_cellmap["LineToNodeX"] = SRCConstants.LineToNodeX;
                CellSRCDictionary.page_cellmap["LineToNodeY"] = SRCConstants.LineToNodeY;
                CellSRCDictionary.page_cellmap["PlaceDepth"] = SRCConstants.PlaceDepth;
                CellSRCDictionary.page_cellmap["PlaceFlip"] = SRCConstants.PlaceFlip;
                CellSRCDictionary.page_cellmap["PlaceStyle"] = SRCConstants.PlaceStyle;
                CellSRCDictionary.page_cellmap["PlowCode"] = SRCConstants.PlowCode;
                CellSRCDictionary.page_cellmap["ResizePage"] = SRCConstants.ResizePage;
                CellSRCDictionary.page_cellmap["RouteStyle"] = SRCConstants.RouteStyle;
                CellSRCDictionary.page_cellmap["AvoidPageBreaks"] = SRCConstants.AvoidPageBreaks;
                CellSRCDictionary.page_cellmap["DrawingResizeType"] = SRCConstants.DrawingResizeType;

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