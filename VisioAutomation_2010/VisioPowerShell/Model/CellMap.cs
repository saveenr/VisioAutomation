using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet;
using VA = VisioAutomation;

namespace VisioPowerShell
{
    public class CellMap
    {
        private Dictionary<string, VA.ShapeSheet.SRC> dic;

        private System.Text.RegularExpressions.Regex regex_cellname;
        private System.Text.RegularExpressions.Regex regex_cellname_wildcard;

        public CellMap()
        {
            this.regex_cellname = new System.Text.RegularExpressions.Regex("^[a-zA-Z]*$");
            this.regex_cellname_wildcard = new System.Text.RegularExpressions.Regex("^[a-zA-Z\\*\\?]*$");
            this.dic = new Dictionary<string, VA.ShapeSheet.SRC>(System.StringComparer.OrdinalIgnoreCase);
        }

        public List<string> GetNames()
        {
            return this.CellNames.ToList();
        }

        public VisioAutomation.ShapeSheet.SRC this[string name]
        {
            get { return this.dic[name]; }
            set
            {
                this.CheckCellName(name);

                if (dic.ContainsKey(name))
                {
                    string msg = string.Format("CellMap already contains a cell called \"{0}\"", name);
                    throw new System.ArgumentOutOfRangeException(msg);
                }
                this.dic[name] = value;
            }
        }

        public Dictionary<string, VA.ShapeSheet.SRC>.KeyCollection CellNames
        {
            get { return this.dic.Keys; }
        }

        public bool IsValidCellName(string name)
        {
            return this.regex_cellname.IsMatch(name);
        }

        public bool IsValidCellNameWildCard(string name)
        {
            return this.regex_cellname_wildcard.IsMatch(name);
        }


        public void CheckCellName(string name)
        {
            if (this.IsValidCellName(name))
            {
                return;
            }

            string msg = string.Format("Cell name \"{0}\" is not valid", name);
            throw new System.ArgumentOutOfRangeException(msg);
        }

        public void CheckCellNameWildcard(string name)
        {
            if (this.IsValidCellNameWildCard(name))
            {
                return;
            }

            string msg = string.Format("Cell name pattern \"{0}\" is not valid", name);
            throw new System.ArgumentOutOfRangeException(msg);
        }

        public IEnumerable<string> ResolveName(string cellname)
        {
            if (cellname.Contains("*") || cellname.Contains("?"))
            {
                this.CheckCellNameWildcard(cellname);
                string pat = "^" + System.Text.RegularExpressions.Regex.Escape(cellname)
                    .Replace(@"\*", ".*").
                    Replace(@"\?", ".") + "$";

                var regex = new System.Text.RegularExpressions.Regex(pat,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                foreach (string k in this.CellNames)
                {
                    if (regex.IsMatch(k))
                    {
                        yield return k;
                    }
                }
            }
            else
            {
                this.CheckCellName(cellname);
                if (!this.dic.ContainsKey(cellname))
                {
                    throw new System.ArgumentException("cellname not defined in map");
                }
                yield return cellname;
            }
        }

        public IEnumerable<string> ResolveNames(string[] cellnames)
        {
            foreach (var name in cellnames)
            {
                foreach (var resolved_name in this.ResolveName(name))
                {
                    yield return resolved_name;
                }
            }

            List<KeyValuePair<string, VA.ShapeSheet.SRC>> pairs = this.dic.ToList();
        }

        public bool ContainsCell(string name)
        {
            return this.dic.ContainsKey(name);
        }

        private static List<VA.ShapeSheet.SRC> shape_cells;
        private static List<VA.ShapeSheet.SRC> page_cells;

        private static CellMap map_name_to_shape_cell;
        private static CellMap map_name_to_page_cell;

        public static List<VA.ShapeSheet.SRC> GetShapeCells()
        {
            if (shape_cells == null)
            {
                shape_cells = new List<SRC>
                {
                    VA.ShapeSheet.SRCConstants.Angle,
                    VA.ShapeSheet.SRCConstants.BeginX,
                    VA.ShapeSheet.SRCConstants.BeginY,
                    VA.ShapeSheet.SRCConstants.CharCase,
                    VA.ShapeSheet.SRCConstants.CharColor,
                    VA.ShapeSheet.SRCConstants.CharColorTrans,
                    VA.ShapeSheet.SRCConstants.CharFont,
                    VA.ShapeSheet.SRCConstants.CharFontScale,
                    VA.ShapeSheet.SRCConstants.CharLetterspace,
                    VA.ShapeSheet.SRCConstants.CharSize,
                    VA.ShapeSheet.SRCConstants.CharStyle,
                    VA.ShapeSheet.SRCConstants.EndX,
                    VA.ShapeSheet.SRCConstants.EndY,
                    VA.ShapeSheet.SRCConstants.FillBkgnd,
                    VA.ShapeSheet.SRCConstants.FillBkgndTrans,
                    VA.ShapeSheet.SRCConstants.FillForegnd,
                    VA.ShapeSheet.SRCConstants.FillForegndTrans,
                    VA.ShapeSheet.SRCConstants.FillPattern,
                    VA.ShapeSheet.SRCConstants.Height,
                    VA.ShapeSheet.SRCConstants.LineCap,
                    VA.ShapeSheet.SRCConstants.LineColor,
                    VA.ShapeSheet.SRCConstants.LinePattern,
                    VA.ShapeSheet.SRCConstants.LineWeight,
                    VA.ShapeSheet.SRCConstants.LockAspect,
                    VA.ShapeSheet.SRCConstants.LockBegin,
                    VA.ShapeSheet.SRCConstants.LockCalcWH,
                    VA.ShapeSheet.SRCConstants.LockCrop,
                    VA.ShapeSheet.SRCConstants.LockCustProp,
                    VA.ShapeSheet.SRCConstants.LockDelete,
                    VA.ShapeSheet.SRCConstants.LockEnd,
                    VA.ShapeSheet.SRCConstants.LockFormat,
                    VA.ShapeSheet.SRCConstants.LockFromGroupFormat,
                    VA.ShapeSheet.SRCConstants.LockGroup,
                    VA.ShapeSheet.SRCConstants.LockHeight,
                    VA.ShapeSheet.SRCConstants.LockMoveX,
                    VA.ShapeSheet.SRCConstants.LockMoveY,
                    VA.ShapeSheet.SRCConstants.LockRotate,
                    VA.ShapeSheet.SRCConstants.LockSelect,
                    VA.ShapeSheet.SRCConstants.LockTextEdit,
                    VA.ShapeSheet.SRCConstants.LockThemeColors,
                    VA.ShapeSheet.SRCConstants.LockThemeEffects,
                    VA.ShapeSheet.SRCConstants.LockVtxEdit,
                    VA.ShapeSheet.SRCConstants.LockWidth,
                    VA.ShapeSheet.SRCConstants.LocPinX,
                    VA.ShapeSheet.SRCConstants.LocPinY,
                    VA.ShapeSheet.SRCConstants.PinX,
                    VA.ShapeSheet.SRCConstants.PinY,
                    VA.ShapeSheet.SRCConstants.Rounding,
                    VA.ShapeSheet.SRCConstants.SelectMode,
                    VA.ShapeSheet.SRCConstants.ShdwBkgnd,
                    VA.ShapeSheet.SRCConstants.ShdwBkgndTrans,
                    VA.ShapeSheet.SRCConstants.ShdwForegnd,
                    VA.ShapeSheet.SRCConstants.ShdwForegndTrans,
                    VA.ShapeSheet.SRCConstants.ShdwObliqueAngle,
                    VA.ShapeSheet.SRCConstants.ShdwOffsetX,
                    VA.ShapeSheet.SRCConstants.ShdwOffsetY,
                    VA.ShapeSheet.SRCConstants.ShdwPattern,
                    VA.ShapeSheet.SRCConstants.ShdwScaleFactor,
                    VA.ShapeSheet.SRCConstants.ShdwType,
                    VA.ShapeSheet.SRCConstants.TxtAngle,
                    VA.ShapeSheet.SRCConstants.TxtHeight,
                    VA.ShapeSheet.SRCConstants.TxtLocPinX,
                    VA.ShapeSheet.SRCConstants.TxtLocPinY,
                    VA.ShapeSheet.SRCConstants.TxtPinX,
                    VA.ShapeSheet.SRCConstants.TxtPinY,
                    VA.ShapeSheet.SRCConstants.TxtWidth,
                    VA.ShapeSheet.SRCConstants.Width,
                    VA.ShapeSheet.SRCConstants.BeginArrow,
                    VA.ShapeSheet.SRCConstants.BeginArrowSize,
                    VA.ShapeSheet.SRCConstants.EndArrow,
                    VA.ShapeSheet.SRCConstants.EndArrowSize,
                    VA.ShapeSheet.SRCConstants.HideText
                };
            }
            return shape_cells;
        }

        public static List<VA.ShapeSheet.SRC> GetPageCells()
        {
            if (page_cells == null)
            {
                page_cells = new List<SRC>
                {
                    VA.ShapeSheet.SRCConstants.PageBottomMargin,
                    VA.ShapeSheet.SRCConstants.PageHeight,
                    VA.ShapeSheet.SRCConstants.PageLeftMargin,
                    VA.ShapeSheet.SRCConstants.PageLineJumpDirX,
                    VA.ShapeSheet.SRCConstants.PageLineJumpDirY,
                    VA.ShapeSheet.SRCConstants.PageRightMargin,
                    VA.ShapeSheet.SRCConstants.PageScale,
                    VA.ShapeSheet.SRCConstants.PageShapeSplit,
                    VA.ShapeSheet.SRCConstants.PageTopMargin,
                    VA.ShapeSheet.SRCConstants.PageWidth,
                    VA.ShapeSheet.SRCConstants.CenterX,
                    VA.ShapeSheet.SRCConstants.CenterY,
                    VA.ShapeSheet.SRCConstants.PaperKind,
                    VA.ShapeSheet.SRCConstants.PrintGrid,
                    VA.ShapeSheet.SRCConstants.PrintPageOrientation,
                    VA.ShapeSheet.SRCConstants.ScaleX,
                    VA.ShapeSheet.SRCConstants.ScaleY,
                    VA.ShapeSheet.SRCConstants.PaperSource,
                    VA.ShapeSheet.SRCConstants.DrawingScale,
                    VA.ShapeSheet.SRCConstants.DrawingScaleType,
                    VA.ShapeSheet.SRCConstants.DrawingSizeType,
                    VA.ShapeSheet.SRCConstants.InhibitSnap,
                    VA.ShapeSheet.SRCConstants.ShdwObliqueAngle,
                    VA.ShapeSheet.SRCConstants.ShdwOffsetX,
                    VA.ShapeSheet.SRCConstants.ShdwOffsetY,
                    VA.ShapeSheet.SRCConstants.ShdwScaleFactor,
                    VA.ShapeSheet.SRCConstants.ShdwType,
                    VA.ShapeSheet.SRCConstants.UIVisibility,
                    VA.ShapeSheet.SRCConstants.XGridDensity,
                    VA.ShapeSheet.SRCConstants.XGridOrigin,
                    VA.ShapeSheet.SRCConstants.XGridSpacing,
                    VA.ShapeSheet.SRCConstants.XRulerDensity,
                    VA.ShapeSheet.SRCConstants.XRulerOrigin,
                    VA.ShapeSheet.SRCConstants.YGridDensity,
                    VA.ShapeSheet.SRCConstants.YGridOrigin,
                    VA.ShapeSheet.SRCConstants.YGridSpacing,
                    VA.ShapeSheet.SRCConstants.YRulerDensity,
                    VA.ShapeSheet.SRCConstants.YRulerOrigin,
                    VA.ShapeSheet.SRCConstants.AvenueSizeX,
                    VA.ShapeSheet.SRCConstants.AvenueSizeY,
                    VA.ShapeSheet.SRCConstants.BlockSizeX,
                    VA.ShapeSheet.SRCConstants.BlockSizeY,
                    VA.ShapeSheet.SRCConstants.CtrlAsInput,
                    VA.ShapeSheet.SRCConstants.DynamicsOff,
                    VA.ShapeSheet.SRCConstants.EnableGrid,
                    VA.ShapeSheet.SRCConstants.LineAdjustFrom,
                    VA.ShapeSheet.SRCConstants.LineAdjustTo,
                    VA.ShapeSheet.SRCConstants.LineJumpCode,
                    VA.ShapeSheet.SRCConstants.LineJumpFactorX,
                    VA.ShapeSheet.SRCConstants.LineJumpFactorY,
                    VA.ShapeSheet.SRCConstants.LineJumpStyle,
                    VA.ShapeSheet.SRCConstants.LineRouteExt,
                    VA.ShapeSheet.SRCConstants.LineToLineX,
                    VA.ShapeSheet.SRCConstants.LineToLineY,
                    VA.ShapeSheet.SRCConstants.LineToNodeX,
                    VA.ShapeSheet.SRCConstants.LineToNodeY,
                    VA.ShapeSheet.SRCConstants.PlaceDepth,
                    VA.ShapeSheet.SRCConstants.PlaceFlip,
                    VA.ShapeSheet.SRCConstants.PlaceStyle,
                    VA.ShapeSheet.SRCConstants.PlowCode,
                    VA.ShapeSheet.SRCConstants.ResizePage,
                    VA.ShapeSheet.SRCConstants.RouteStyle,
                    VA.ShapeSheet.SRCConstants.AvoidPageBreaks,
                    VA.ShapeSheet.SRCConstants.DrawingResizeType
                };
            }
            return page_cells;
        }


        public static CellMap GetShapeCellDictionary()
        {
            if (map_name_to_shape_cell == null)
            {
                map_name_to_shape_cell = new CellMap();
                map_name_to_shape_cell["Angle"] = VA.ShapeSheet.SRCConstants.Angle;
                map_name_to_shape_cell["BeginX"] = VA.ShapeSheet.SRCConstants.BeginX;
                map_name_to_shape_cell["BeginY"] = VA.ShapeSheet.SRCConstants.BeginY;
                map_name_to_shape_cell["CharCase"] = VA.ShapeSheet.SRCConstants.CharCase;
                map_name_to_shape_cell["CharColor"] = VA.ShapeSheet.SRCConstants.CharColor;
                map_name_to_shape_cell["CharColorTransparency"] = VA.ShapeSheet.SRCConstants.CharColorTrans;
                map_name_to_shape_cell["CharFont"] = VA.ShapeSheet.SRCConstants.CharFont;
                map_name_to_shape_cell["CharFontScale"] = VA.ShapeSheet.SRCConstants.CharFontScale;
                map_name_to_shape_cell["CharLetterspace"] = VA.ShapeSheet.SRCConstants.CharLetterspace;
                map_name_to_shape_cell["CharSize"] = VA.ShapeSheet.SRCConstants.CharSize;
                map_name_to_shape_cell["CharStyle"] = VA.ShapeSheet.SRCConstants.CharStyle;
                map_name_to_shape_cell["EndX"] = VA.ShapeSheet.SRCConstants.EndX;
                map_name_to_shape_cell["EndY"] = VA.ShapeSheet.SRCConstants.EndY;
                map_name_to_shape_cell["FillBkgnd"] = VA.ShapeSheet.SRCConstants.FillBkgnd;
                map_name_to_shape_cell["FillBkgndTrans"] = VA.ShapeSheet.SRCConstants.FillBkgndTrans;
                map_name_to_shape_cell["FillForegnd"] = VA.ShapeSheet.SRCConstants.FillForegnd;
                map_name_to_shape_cell["FillForegndTrans"] = VA.ShapeSheet.SRCConstants.FillForegndTrans;
                map_name_to_shape_cell["FillPattern"] = VA.ShapeSheet.SRCConstants.FillPattern;
                map_name_to_shape_cell["Height"] = VA.ShapeSheet.SRCConstants.Height;
                map_name_to_shape_cell["LineCap"] = VA.ShapeSheet.SRCConstants.LineCap;
                map_name_to_shape_cell["LineColor"] = VA.ShapeSheet.SRCConstants.LineColor;
                map_name_to_shape_cell["LinePattern"] = VA.ShapeSheet.SRCConstants.LinePattern;
                map_name_to_shape_cell["LineWeight"] = VA.ShapeSheet.SRCConstants.LineWeight;
                map_name_to_shape_cell["LockAspect"] = VA.ShapeSheet.SRCConstants.LockAspect;
                map_name_to_shape_cell["LockBegin"] = VA.ShapeSheet.SRCConstants.LockBegin;
                map_name_to_shape_cell["LockCalcWH"] = VA.ShapeSheet.SRCConstants.LockCalcWH;
                map_name_to_shape_cell["LockCrop"] = VA.ShapeSheet.SRCConstants.LockCrop;
                map_name_to_shape_cell["LockCustProp"] = VA.ShapeSheet.SRCConstants.LockCustProp;
                map_name_to_shape_cell["LockDelete"] = VA.ShapeSheet.SRCConstants.LockDelete;
                map_name_to_shape_cell["LockEnd"] = VA.ShapeSheet.SRCConstants.LockEnd;
                map_name_to_shape_cell["LockFormat"] = VA.ShapeSheet.SRCConstants.LockFormat;
                map_name_to_shape_cell["LockFromGroupFormat"] = VA.ShapeSheet.SRCConstants.LockFromGroupFormat;
                map_name_to_shape_cell["LockGroup"] = VA.ShapeSheet.SRCConstants.LockGroup;
                map_name_to_shape_cell["LockHeight"] = VA.ShapeSheet.SRCConstants.LockHeight;
                map_name_to_shape_cell["LockMoveX"] = VA.ShapeSheet.SRCConstants.LockMoveX;
                map_name_to_shape_cell["LockMoveY"] = VA.ShapeSheet.SRCConstants.LockMoveY;
                map_name_to_shape_cell["LockRotate"] = VA.ShapeSheet.SRCConstants.LockRotate;
                map_name_to_shape_cell["LockSelect"] = VA.ShapeSheet.SRCConstants.LockSelect;
                map_name_to_shape_cell["LockTextEdit"] = VA.ShapeSheet.SRCConstants.LockTextEdit;
                map_name_to_shape_cell["LockThemeColors"] = VA.ShapeSheet.SRCConstants.LockThemeColors;
                map_name_to_shape_cell["LockThemeEffects"] = VA.ShapeSheet.SRCConstants.LockThemeEffects;
                map_name_to_shape_cell["LockVtxEdit"] = VA.ShapeSheet.SRCConstants.LockVtxEdit;
                map_name_to_shape_cell["LockWidth"] = VA.ShapeSheet.SRCConstants.LockWidth;
                map_name_to_shape_cell["LocPinX"] = VA.ShapeSheet.SRCConstants.LocPinX;
                map_name_to_shape_cell["LocPinY"] = VA.ShapeSheet.SRCConstants.LocPinY;
                map_name_to_shape_cell["PinX"] = VA.ShapeSheet.SRCConstants.PinX;
                map_name_to_shape_cell["PinY"] = VA.ShapeSheet.SRCConstants.PinY;
                map_name_to_shape_cell["Rounding"] = VA.ShapeSheet.SRCConstants.Rounding;
                map_name_to_shape_cell["SelectMode"] = VA.ShapeSheet.SRCConstants.SelectMode;
                map_name_to_shape_cell["ShdwBkgnd"] = VA.ShapeSheet.SRCConstants.ShdwBkgnd;
                map_name_to_shape_cell["ShdwBkgndTrans"] = VA.ShapeSheet.SRCConstants.ShdwBkgndTrans;
                map_name_to_shape_cell["ShdwForegnd"] = VA.ShapeSheet.SRCConstants.ShdwForegnd;
                map_name_to_shape_cell["ShdwForegndTrans"] = VA.ShapeSheet.SRCConstants.ShdwForegndTrans;
                map_name_to_shape_cell["ShdwObliqueAngle"] = VA.ShapeSheet.SRCConstants.ShdwObliqueAngle;
                map_name_to_shape_cell["ShdwOffsetX"] = VA.ShapeSheet.SRCConstants.ShdwOffsetX;
                map_name_to_shape_cell["ShdwOffsetY"] = VA.ShapeSheet.SRCConstants.ShdwOffsetY;
                map_name_to_shape_cell["ShdwPattern"] = VA.ShapeSheet.SRCConstants.ShdwPattern;
                map_name_to_shape_cell["ShdwScaleFactor"] = VA.ShapeSheet.SRCConstants.ShdwScaleFactor;
                map_name_to_shape_cell["ShdwType"] = VA.ShapeSheet.SRCConstants.ShdwType;
                map_name_to_shape_cell["TxtAngle"] = VA.ShapeSheet.SRCConstants.TxtAngle;
                map_name_to_shape_cell["TxtHeight"] = VA.ShapeSheet.SRCConstants.TxtHeight;
                map_name_to_shape_cell["TxtLocPinX"] = VA.ShapeSheet.SRCConstants.TxtLocPinX;
                map_name_to_shape_cell["TxtLocPinY"] = VA.ShapeSheet.SRCConstants.TxtLocPinY;
                map_name_to_shape_cell["TxtPinX"] = VA.ShapeSheet.SRCConstants.TxtPinX;
                map_name_to_shape_cell["TxtPinY"] = VA.ShapeSheet.SRCConstants.TxtPinY;
                map_name_to_shape_cell["TxtWidth"] = VA.ShapeSheet.SRCConstants.TxtWidth;
                map_name_to_shape_cell["Width"] = VA.ShapeSheet.SRCConstants.Width;  

            }
            return map_name_to_shape_cell;
        }

        public static CellMap GetPageCellDictionary()
        {
            if (map_name_to_page_cell == null)
            {
                map_name_to_page_cell = new CellMap();
                                    map_name_to_page_cell["PageBottomMargin"] = VA.ShapeSheet.SRCConstants.PageBottomMargin;   
                    map_name_to_page_cell["PageHeight"] = VA.ShapeSheet.SRCConstants.PageHeight;   
                    map_name_to_page_cell["PageLeftMargin"] = VA.ShapeSheet.SRCConstants.PageLeftMargin;   
                    map_name_to_page_cell["PageLineJumpDirX"] = VA.ShapeSheet.SRCConstants.PageLineJumpDirX;   
                    map_name_to_page_cell["PageLineJumpDirY"] = VA.ShapeSheet.SRCConstants.PageLineJumpDirY;   

                    map_name_to_page_cell["PageRightMargin"] = VA.ShapeSheet.SRCConstants.PageRightMargin;   
                    map_name_to_page_cell["PageScale"] = VA.ShapeSheet.SRCConstants.PageScale;   
                    map_name_to_page_cell["PageShapeSplit"] = VA.ShapeSheet.SRCConstants.PageShapeSplit;   
                    map_name_to_page_cell["PageTopMargin"] = VA.ShapeSheet.SRCConstants.PageTopMargin;   
                    map_name_to_page_cell["PageWidth"] = VA.ShapeSheet.SRCConstants.PageWidth;   
                    map_name_to_page_cell["CenterX"] = VA.ShapeSheet.SRCConstants.CenterX;   
                    map_name_to_page_cell["CenterY"] = VA.ShapeSheet.SRCConstants.CenterY;   
                    map_name_to_page_cell["PaperKind"] = VA.ShapeSheet.SRCConstants.PaperKind;   
                    map_name_to_page_cell["PrintGrid"] = VA.ShapeSheet.SRCConstants.PrintGrid;   
                    map_name_to_page_cell["PrintPageOrientation"] = VA.ShapeSheet.SRCConstants.PrintPageOrientation;   
                    map_name_to_page_cell["ScaleX"] = VA.ShapeSheet.SRCConstants.ScaleX;   
                    map_name_to_page_cell["ScaleY"] = VA.ShapeSheet.SRCConstants.ScaleY;   
                    map_name_to_page_cell["PaperSource"] = VA.ShapeSheet.SRCConstants.PaperSource;   
                    map_name_to_page_cell["DrawingScale"] = VA.ShapeSheet.SRCConstants.DrawingScale;   
                    map_name_to_page_cell["DrawingScaleType"] = VA.ShapeSheet.SRCConstants.DrawingScaleType;   
                    map_name_to_page_cell["DrawingSizeType"] = VA.ShapeSheet.SRCConstants.DrawingSizeType;   
                    map_name_to_page_cell["InhibitSnap"] = VA.ShapeSheet.SRCConstants.InhibitSnap;   
                    map_name_to_page_cell["ShdwObliqueAngle"] = VA.ShapeSheet.SRCConstants.ShdwObliqueAngle;   
                    map_name_to_page_cell["ShdwOffsetX"] = VA.ShapeSheet.SRCConstants.ShdwOffsetX;   
                    map_name_to_page_cell["ShdwOffsetY"] = VA.ShapeSheet.SRCConstants.ShdwOffsetY;   
                    map_name_to_page_cell["ShdwScaleFactor"] = VA.ShapeSheet.SRCConstants.ShdwScaleFactor;   
                    map_name_to_page_cell["ShdwType"] = VA.ShapeSheet.SRCConstants.ShdwType;   
                    map_name_to_page_cell["UIVisibility"] = VA.ShapeSheet.SRCConstants.UIVisibility;   
                    map_name_to_page_cell["XGridDensity"] = VA.ShapeSheet.SRCConstants.XGridDensity;   
                    map_name_to_page_cell["XGridOrigin"] = VA.ShapeSheet.SRCConstants.XGridOrigin;   
                    map_name_to_page_cell["XGridSpacing"] = VA.ShapeSheet.SRCConstants.XGridSpacing;   
                    map_name_to_page_cell["XRulerDensity"] = VA.ShapeSheet.SRCConstants.XRulerDensity;   
                    map_name_to_page_cell["XRulerOrigin"] = VA.ShapeSheet.SRCConstants.XRulerOrigin;   
                    map_name_to_page_cell["YGridDensity"] = VA.ShapeSheet.SRCConstants.YGridDensity;   
                    map_name_to_page_cell["YGridOrigin"] = VA.ShapeSheet.SRCConstants.YGridOrigin;   
                    map_name_to_page_cell["YGridSpacing"] = VA.ShapeSheet.SRCConstants.YGridSpacing;   
                    map_name_to_page_cell["YRulerDensity"] = VA.ShapeSheet.SRCConstants.YRulerDensity;   
                    map_name_to_page_cell["YRulerOrigin"] = VA.ShapeSheet.SRCConstants.YRulerOrigin;   
                    map_name_to_page_cell["AvenueSizeX"] = VA.ShapeSheet.SRCConstants.AvenueSizeX;   
                    map_name_to_page_cell["AvenueSizeY"] = VA.ShapeSheet.SRCConstants.AvenueSizeY;   
                    map_name_to_page_cell["BlockSizeX"] = VA.ShapeSheet.SRCConstants.BlockSizeX;   
                    map_name_to_page_cell["BlockSizeY"] = VA.ShapeSheet.SRCConstants.BlockSizeY;   
                    map_name_to_page_cell["CtrlAsInput"] = VA.ShapeSheet.SRCConstants.CtrlAsInput;   
                    map_name_to_page_cell["DynamicsOff"] = VA.ShapeSheet.SRCConstants.DynamicsOff;   
                    map_name_to_page_cell["EnableGrid"] = VA.ShapeSheet.SRCConstants.EnableGrid;   
                    map_name_to_page_cell["LineAdjustFrom"] = VA.ShapeSheet.SRCConstants.LineAdjustFrom;   
                    map_name_to_page_cell["LineAdjustTo"] = VA.ShapeSheet.SRCConstants.LineAdjustTo;   
                    map_name_to_page_cell["LineJumpCode"] = VA.ShapeSheet.SRCConstants.LineJumpCode;   
                    map_name_to_page_cell["LineJumpFactorX"] = VA.ShapeSheet.SRCConstants.LineJumpFactorX;   
                    map_name_to_page_cell["LineJumpFactorY"] = VA.ShapeSheet.SRCConstants.LineJumpFactorY;   
                    map_name_to_page_cell["LineJumpStyle"] = VA.ShapeSheet.SRCConstants.LineJumpStyle;   
                    map_name_to_page_cell["LineRouteExt"] = VA.ShapeSheet.SRCConstants.LineRouteExt;   
                    map_name_to_page_cell["LineToLineX"] = VA.ShapeSheet.SRCConstants.LineToLineX;   
                    map_name_to_page_cell["LineToLineY"] = VA.ShapeSheet.SRCConstants.LineToLineY;   
                    map_name_to_page_cell["LineToNodeX"] = VA.ShapeSheet.SRCConstants.LineToNodeX;   
                    map_name_to_page_cell["LineToNodeY"] = VA.ShapeSheet.SRCConstants.LineToNodeY;   
                    map_name_to_page_cell["PlaceDepth"] = VA.ShapeSheet.SRCConstants.PlaceDepth;   
                    map_name_to_page_cell["PlaceFlip"] = VA.ShapeSheet.SRCConstants.PlaceFlip;   
                    map_name_to_page_cell["PlaceStyle"] = VA.ShapeSheet.SRCConstants.PlaceStyle;   
                    map_name_to_page_cell["PlowCode"] = VA.ShapeSheet.SRCConstants.PlowCode;   
                    map_name_to_page_cell["ResizePage"] = VA.ShapeSheet.SRCConstants.ResizePage;   
                    map_name_to_page_cell["RouteStyle"] = VA.ShapeSheet.SRCConstants.RouteStyle;   
                    map_name_to_page_cell["AvoidPageBreaks"] = VA.ShapeSheet.SRCConstants.AvoidPageBreaks;   
                    map_name_to_page_cell["DrawingResizeType"] = VA.ShapeSheet.SRCConstants.DrawingResizeType;  

            }
            return map_name_to_page_cell;
        }

        public VisioAutomation.ShapeSheet.Query.CellQuery CreateQueryFromCellNames(string[] Cells)
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
                    query.AddCell(dic[resolved_cellname], resolved_cellname);
                }
            }
            return query;
        }
    }
}