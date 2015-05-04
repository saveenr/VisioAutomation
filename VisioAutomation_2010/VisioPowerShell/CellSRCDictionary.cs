using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;

namespace VisioPowerShell
{
    public class CellSRCDictionary : CellNameDictionary<VA.ShapeSheet.SRC>
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
                CellSRCDictionary.shape_cellmap["Angle"] = VA.ShapeSheet.SRCConstants.Angle;
                CellSRCDictionary.shape_cellmap["BeginX"] = VA.ShapeSheet.SRCConstants.BeginX;
                CellSRCDictionary.shape_cellmap["BeginY"] = VA.ShapeSheet.SRCConstants.BeginY;
                CellSRCDictionary.shape_cellmap["CharCase"] = VA.ShapeSheet.SRCConstants.CharCase;
                CellSRCDictionary.shape_cellmap["CharColor"] = VA.ShapeSheet.SRCConstants.CharColor;
                CellSRCDictionary.shape_cellmap["CharColorTrans"] = VA.ShapeSheet.SRCConstants.CharColorTrans;
                CellSRCDictionary.shape_cellmap["CharFont"] = VA.ShapeSheet.SRCConstants.CharFont;
                CellSRCDictionary.shape_cellmap["CharFontScale"] = VA.ShapeSheet.SRCConstants.CharFontScale;
                CellSRCDictionary.shape_cellmap["CharLetterspace"] = VA.ShapeSheet.SRCConstants.CharLetterspace;
                CellSRCDictionary.shape_cellmap["CharSize"] = VA.ShapeSheet.SRCConstants.CharSize;
                CellSRCDictionary.shape_cellmap["CharStyle"] = VA.ShapeSheet.SRCConstants.CharStyle;
                CellSRCDictionary.shape_cellmap["EndX"] = VA.ShapeSheet.SRCConstants.EndX;
                CellSRCDictionary.shape_cellmap["EndY"] = VA.ShapeSheet.SRCConstants.EndY;
                CellSRCDictionary.shape_cellmap["FillBkgnd"] = VA.ShapeSheet.SRCConstants.FillBkgnd;
                CellSRCDictionary.shape_cellmap["FillBkgndTrans"] = VA.ShapeSheet.SRCConstants.FillBkgndTrans;
                CellSRCDictionary.shape_cellmap["FillForegnd"] = VA.ShapeSheet.SRCConstants.FillForegnd;
                CellSRCDictionary.shape_cellmap["FillForegndTrans"] = VA.ShapeSheet.SRCConstants.FillForegndTrans;
                CellSRCDictionary.shape_cellmap["FillPattern"] = VA.ShapeSheet.SRCConstants.FillPattern;
                CellSRCDictionary.shape_cellmap["Height"] = VA.ShapeSheet.SRCConstants.Height;
                CellSRCDictionary.shape_cellmap["LineCap"] = VA.ShapeSheet.SRCConstants.LineCap;
                CellSRCDictionary.shape_cellmap["LineColor"] = VA.ShapeSheet.SRCConstants.LineColor;
                CellSRCDictionary.shape_cellmap["LinePattern"] = VA.ShapeSheet.SRCConstants.LinePattern;
                CellSRCDictionary.shape_cellmap["LineWeight"] = VA.ShapeSheet.SRCConstants.LineWeight;
                CellSRCDictionary.shape_cellmap["LockAspect"] = VA.ShapeSheet.SRCConstants.LockAspect;
                CellSRCDictionary.shape_cellmap["LockBegin"] = VA.ShapeSheet.SRCConstants.LockBegin;
                CellSRCDictionary.shape_cellmap["LockCalcWH"] = VA.ShapeSheet.SRCConstants.LockCalcWH;
                CellSRCDictionary.shape_cellmap["LockCrop"] = VA.ShapeSheet.SRCConstants.LockCrop;
                CellSRCDictionary.shape_cellmap["LockCustProp"] = VA.ShapeSheet.SRCConstants.LockCustProp;
                CellSRCDictionary.shape_cellmap["LockDelete"] = VA.ShapeSheet.SRCConstants.LockDelete;
                CellSRCDictionary.shape_cellmap["LockEnd"] = VA.ShapeSheet.SRCConstants.LockEnd;
                CellSRCDictionary.shape_cellmap["LockFormat"] = VA.ShapeSheet.SRCConstants.LockFormat;
                CellSRCDictionary.shape_cellmap["LockFromGroupFormat"] = VA.ShapeSheet.SRCConstants.LockFromGroupFormat;
                CellSRCDictionary.shape_cellmap["LockGroup"] = VA.ShapeSheet.SRCConstants.LockGroup;
                CellSRCDictionary.shape_cellmap["LockHeight"] = VA.ShapeSheet.SRCConstants.LockHeight;
                CellSRCDictionary.shape_cellmap["LockMoveX"] = VA.ShapeSheet.SRCConstants.LockMoveX;
                CellSRCDictionary.shape_cellmap["LockMoveY"] = VA.ShapeSheet.SRCConstants.LockMoveY;
                CellSRCDictionary.shape_cellmap["LockRotate"] = VA.ShapeSheet.SRCConstants.LockRotate;
                CellSRCDictionary.shape_cellmap["LockSelect"] = VA.ShapeSheet.SRCConstants.LockSelect;
                CellSRCDictionary.shape_cellmap["LockTextEdit"] = VA.ShapeSheet.SRCConstants.LockTextEdit;
                CellSRCDictionary.shape_cellmap["LockThemeColors"] = VA.ShapeSheet.SRCConstants.LockThemeColors;
                CellSRCDictionary.shape_cellmap["LockThemeEffects"] = VA.ShapeSheet.SRCConstants.LockThemeEffects;
                CellSRCDictionary.shape_cellmap["LockVtxEdit"] = VA.ShapeSheet.SRCConstants.LockVtxEdit;
                CellSRCDictionary.shape_cellmap["LockWidth"] = VA.ShapeSheet.SRCConstants.LockWidth;
                CellSRCDictionary.shape_cellmap["LocPinX"] = VA.ShapeSheet.SRCConstants.LocPinX;
                CellSRCDictionary.shape_cellmap["LocPinY"] = VA.ShapeSheet.SRCConstants.LocPinY;
                CellSRCDictionary.shape_cellmap["PinX"] = VA.ShapeSheet.SRCConstants.PinX;
                CellSRCDictionary.shape_cellmap["PinY"] = VA.ShapeSheet.SRCConstants.PinY;
                CellSRCDictionary.shape_cellmap["Rounding"] = VA.ShapeSheet.SRCConstants.Rounding;
                CellSRCDictionary.shape_cellmap["SelectMode"] = VA.ShapeSheet.SRCConstants.SelectMode;
                CellSRCDictionary.shape_cellmap["ShdwBkgnd"] = VA.ShapeSheet.SRCConstants.ShdwBkgnd;
                CellSRCDictionary.shape_cellmap["ShdwBkgndTrans"] = VA.ShapeSheet.SRCConstants.ShdwBkgndTrans;
                CellSRCDictionary.shape_cellmap["ShdwForegnd"] = VA.ShapeSheet.SRCConstants.ShdwForegnd;
                CellSRCDictionary.shape_cellmap["ShdwForegndTrans"] = VA.ShapeSheet.SRCConstants.ShdwForegndTrans;
                CellSRCDictionary.shape_cellmap["ShdwObliqueAngle"] = VA.ShapeSheet.SRCConstants.ShdwObliqueAngle;
                CellSRCDictionary.shape_cellmap["ShdwOffsetX"] = VA.ShapeSheet.SRCConstants.ShdwOffsetX;
                CellSRCDictionary.shape_cellmap["ShdwOffsetY"] = VA.ShapeSheet.SRCConstants.ShdwOffsetY;
                CellSRCDictionary.shape_cellmap["ShdwPattern"] = VA.ShapeSheet.SRCConstants.ShdwPattern;
                CellSRCDictionary.shape_cellmap["ShdwScaleFactor"] = VA.ShapeSheet.SRCConstants.ShdwScaleFactor;
                CellSRCDictionary.shape_cellmap["ShdwType"] = VA.ShapeSheet.SRCConstants.ShdwType;
                CellSRCDictionary.shape_cellmap["TxtAngle"] = VA.ShapeSheet.SRCConstants.TxtAngle;
                CellSRCDictionary.shape_cellmap["TxtHeight"] = VA.ShapeSheet.SRCConstants.TxtHeight;
                CellSRCDictionary.shape_cellmap["TxtLocPinX"] = VA.ShapeSheet.SRCConstants.TxtLocPinX;
                CellSRCDictionary.shape_cellmap["TxtLocPinY"] = VA.ShapeSheet.SRCConstants.TxtLocPinY;
                CellSRCDictionary.shape_cellmap["TxtPinX"] = VA.ShapeSheet.SRCConstants.TxtPinX;
                CellSRCDictionary.shape_cellmap["TxtPinY"] = VA.ShapeSheet.SRCConstants.TxtPinY;
                CellSRCDictionary.shape_cellmap["TxtWidth"] = VA.ShapeSheet.SRCConstants.TxtWidth;
                CellSRCDictionary.shape_cellmap["Width"] = VA.ShapeSheet.SRCConstants.Width;

            }
            return CellSRCDictionary.shape_cellmap;
        }

        public static CellSRCDictionary GetCellMapForPages()
        {
            if (CellSRCDictionary.page_cellmap == null)
            {
                CellSRCDictionary.page_cellmap = new CellSRCDictionary();
                CellSRCDictionary.page_cellmap["PageBottomMargin"] = VA.ShapeSheet.SRCConstants.PageBottomMargin;
                CellSRCDictionary.page_cellmap["PageHeight"] = VA.ShapeSheet.SRCConstants.PageHeight;
                CellSRCDictionary.page_cellmap["PageLeftMargin"] = VA.ShapeSheet.SRCConstants.PageLeftMargin;
                CellSRCDictionary.page_cellmap["PageLineJumpDirX"] = VA.ShapeSheet.SRCConstants.PageLineJumpDirX;
                CellSRCDictionary.page_cellmap["PageLineJumpDirY"] = VA.ShapeSheet.SRCConstants.PageLineJumpDirY;

                CellSRCDictionary.page_cellmap["PageRightMargin"] = VA.ShapeSheet.SRCConstants.PageRightMargin;
                CellSRCDictionary.page_cellmap["PageScale"] = VA.ShapeSheet.SRCConstants.PageScale;
                CellSRCDictionary.page_cellmap["PageShapeSplit"] = VA.ShapeSheet.SRCConstants.PageShapeSplit;
                CellSRCDictionary.page_cellmap["PageTopMargin"] = VA.ShapeSheet.SRCConstants.PageTopMargin;
                CellSRCDictionary.page_cellmap["PageWidth"] = VA.ShapeSheet.SRCConstants.PageWidth;
                CellSRCDictionary.page_cellmap["CenterX"] = VA.ShapeSheet.SRCConstants.CenterX;
                CellSRCDictionary.page_cellmap["CenterY"] = VA.ShapeSheet.SRCConstants.CenterY;
                CellSRCDictionary.page_cellmap["PaperKind"] = VA.ShapeSheet.SRCConstants.PaperKind;
                CellSRCDictionary.page_cellmap["PrintGrid"] = VA.ShapeSheet.SRCConstants.PrintGrid;
                CellSRCDictionary.page_cellmap["PrintPageOrientation"] = VA.ShapeSheet.SRCConstants.PrintPageOrientation;
                CellSRCDictionary.page_cellmap["ScaleX"] = VA.ShapeSheet.SRCConstants.ScaleX;
                CellSRCDictionary.page_cellmap["ScaleY"] = VA.ShapeSheet.SRCConstants.ScaleY;
                CellSRCDictionary.page_cellmap["PaperSource"] = VA.ShapeSheet.SRCConstants.PaperSource;
                CellSRCDictionary.page_cellmap["DrawingScale"] = VA.ShapeSheet.SRCConstants.DrawingScale;
                CellSRCDictionary.page_cellmap["DrawingScaleType"] = VA.ShapeSheet.SRCConstants.DrawingScaleType;
                CellSRCDictionary.page_cellmap["DrawingSizeType"] = VA.ShapeSheet.SRCConstants.DrawingSizeType;
                CellSRCDictionary.page_cellmap["InhibitSnap"] = VA.ShapeSheet.SRCConstants.InhibitSnap;
                CellSRCDictionary.page_cellmap["ShdwObliqueAngle"] = VA.ShapeSheet.SRCConstants.ShdwObliqueAngle;
                CellSRCDictionary.page_cellmap["ShdwOffsetX"] = VA.ShapeSheet.SRCConstants.ShdwOffsetX;
                CellSRCDictionary.page_cellmap["ShdwOffsetY"] = VA.ShapeSheet.SRCConstants.ShdwOffsetY;
                CellSRCDictionary.page_cellmap["ShdwScaleFactor"] = VA.ShapeSheet.SRCConstants.ShdwScaleFactor;
                CellSRCDictionary.page_cellmap["ShdwType"] = VA.ShapeSheet.SRCConstants.ShdwType;
                CellSRCDictionary.page_cellmap["UIVisibility"] = VA.ShapeSheet.SRCConstants.UIVisibility;
                CellSRCDictionary.page_cellmap["XGridDensity"] = VA.ShapeSheet.SRCConstants.XGridDensity;
                CellSRCDictionary.page_cellmap["XGridOrigin"] = VA.ShapeSheet.SRCConstants.XGridOrigin;
                CellSRCDictionary.page_cellmap["XGridSpacing"] = VA.ShapeSheet.SRCConstants.XGridSpacing;
                CellSRCDictionary.page_cellmap["XRulerDensity"] = VA.ShapeSheet.SRCConstants.XRulerDensity;
                CellSRCDictionary.page_cellmap["XRulerOrigin"] = VA.ShapeSheet.SRCConstants.XRulerOrigin;
                CellSRCDictionary.page_cellmap["YGridDensity"] = VA.ShapeSheet.SRCConstants.YGridDensity;
                CellSRCDictionary.page_cellmap["YGridOrigin"] = VA.ShapeSheet.SRCConstants.YGridOrigin;
                CellSRCDictionary.page_cellmap["YGridSpacing"] = VA.ShapeSheet.SRCConstants.YGridSpacing;
                CellSRCDictionary.page_cellmap["YRulerDensity"] = VA.ShapeSheet.SRCConstants.YRulerDensity;
                CellSRCDictionary.page_cellmap["YRulerOrigin"] = VA.ShapeSheet.SRCConstants.YRulerOrigin;
                CellSRCDictionary.page_cellmap["AvenueSizeX"] = VA.ShapeSheet.SRCConstants.AvenueSizeX;
                CellSRCDictionary.page_cellmap["AvenueSizeY"] = VA.ShapeSheet.SRCConstants.AvenueSizeY;
                CellSRCDictionary.page_cellmap["BlockSizeX"] = VA.ShapeSheet.SRCConstants.BlockSizeX;
                CellSRCDictionary.page_cellmap["BlockSizeY"] = VA.ShapeSheet.SRCConstants.BlockSizeY;
                CellSRCDictionary.page_cellmap["CtrlAsInput"] = VA.ShapeSheet.SRCConstants.CtrlAsInput;
                CellSRCDictionary.page_cellmap["DynamicsOff"] = VA.ShapeSheet.SRCConstants.DynamicsOff;
                CellSRCDictionary.page_cellmap["EnableGrid"] = VA.ShapeSheet.SRCConstants.EnableGrid;
                CellSRCDictionary.page_cellmap["LineAdjustFrom"] = VA.ShapeSheet.SRCConstants.LineAdjustFrom;
                CellSRCDictionary.page_cellmap["LineAdjustTo"] = VA.ShapeSheet.SRCConstants.LineAdjustTo;
                CellSRCDictionary.page_cellmap["LineJumpCode"] = VA.ShapeSheet.SRCConstants.LineJumpCode;
                CellSRCDictionary.page_cellmap["LineJumpFactorX"] = VA.ShapeSheet.SRCConstants.LineJumpFactorX;
                CellSRCDictionary.page_cellmap["LineJumpFactorY"] = VA.ShapeSheet.SRCConstants.LineJumpFactorY;
                CellSRCDictionary.page_cellmap["LineJumpStyle"] = VA.ShapeSheet.SRCConstants.LineJumpStyle;
                CellSRCDictionary.page_cellmap["LineRouteExt"] = VA.ShapeSheet.SRCConstants.LineRouteExt;
                CellSRCDictionary.page_cellmap["LineToLineX"] = VA.ShapeSheet.SRCConstants.LineToLineX;
                CellSRCDictionary.page_cellmap["LineToLineY"] = VA.ShapeSheet.SRCConstants.LineToLineY;
                CellSRCDictionary.page_cellmap["LineToNodeX"] = VA.ShapeSheet.SRCConstants.LineToNodeX;
                CellSRCDictionary.page_cellmap["LineToNodeY"] = VA.ShapeSheet.SRCConstants.LineToNodeY;
                CellSRCDictionary.page_cellmap["PlaceDepth"] = VA.ShapeSheet.SRCConstants.PlaceDepth;
                CellSRCDictionary.page_cellmap["PlaceFlip"] = VA.ShapeSheet.SRCConstants.PlaceFlip;
                CellSRCDictionary.page_cellmap["PlaceStyle"] = VA.ShapeSheet.SRCConstants.PlaceStyle;
                CellSRCDictionary.page_cellmap["PlowCode"] = VA.ShapeSheet.SRCConstants.PlowCode;
                CellSRCDictionary.page_cellmap["ResizePage"] = VA.ShapeSheet.SRCConstants.ResizePage;
                CellSRCDictionary.page_cellmap["RouteStyle"] = VA.ShapeSheet.SRCConstants.RouteStyle;
                CellSRCDictionary.page_cellmap["AvoidPageBreaks"] = VA.ShapeSheet.SRCConstants.AvoidPageBreaks;
                CellSRCDictionary.page_cellmap["DrawingResizeType"] = VA.ShapeSheet.SRCConstants.DrawingResizeType;

            }
            return CellSRCDictionary.page_cellmap;
        }

        public VisioAutomation.ShapeSheet.Query.CellQuery CreateQueryFromCellNames(IList<string> Cells)
        {
            var invalid_names = Cells.Where(cellname => !this.ContainsCell(cellname)).ToList();
            if (invalid_names.Count > 0)
            {
                string msg = "Invalid cell names: " + string.Join(",", invalid_names);
                throw new System.ArgumentException(msg);
            }

            var query = new VisioAutomation.ShapeSheet.Query.CellQuery();

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