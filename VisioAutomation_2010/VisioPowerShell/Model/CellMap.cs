using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;

namespace VisioPowerShell
{
    /// <summary>
    /// CellMap is simply a dictionary of cell names (strings) to cell SRC values.
    /// It's advantage over a normal dictionary is that it supports looking up values when
    /// wildcards are used for the keys.
    /// </summary>
    public class CellMap
    {
        private Dictionary<string, VA.ShapeSheet.SRC> name_to_src;

        private System.Text.RegularExpressions.Regex regex_cellname;
        private System.Text.RegularExpressions.Regex regex_cellname_wildcard;

        private static CellMap shape_cellmap;
        private static CellMap page_cellmap;

        public CellMap()
        {
            this.regex_cellname = new System.Text.RegularExpressions.Regex("^[a-zA-Z]*$");
            this.regex_cellname_wildcard = new System.Text.RegularExpressions.Regex("^[a-zA-Z\\*\\?]*$");
            this.name_to_src = new Dictionary<string, VA.ShapeSheet.SRC>(System.StringComparer.OrdinalIgnoreCase);
        }

        public List<string> GetNames()
        {
            return this.CellNames.ToList();
        }

        public VisioAutomation.ShapeSheet.SRC this[string name]
        {
            get
            {
                return this.name_to_src[name];
            }
            set
            {
                this.CheckCellName(name);

                if (name_to_src.ContainsKey(name))
                {
                    string msg = string.Format("CellMap already contains a cell called \"{0}\"", name);
                    throw new System.ArgumentOutOfRangeException(msg);
                }

                this.name_to_src[name] = value;
            }
        }

        public Dictionary<string, VA.ShapeSheet.SRC>.KeyCollection CellNames
        {
            get
            {
                return this.name_to_src.Keys;
            }
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

            string msg = string.Format("Cell name wildcard pattern \"{0}\" is not valid", name);
            throw new System.ArgumentException(msg,"name");
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
                if (this.name_to_src.ContainsKey(cellname))
                {
                    // found the exact cell name, yield it
                    yield return cellname;                    
                }
                else 
                {
                    // Coudn't find the exact cell name, yield nothing
                    yield break;
                }
            }
        }

        public IEnumerable<string> ResolveNames(IEnumerable<string> cellnames)
        {
            foreach (var name in cellnames)
            {
                foreach (var resolved_name in this.ResolveName(name))
                {
                    yield return resolved_name;
                }
            }
        }

        public bool ContainsCell(string name)
        {
            return this.name_to_src.ContainsKey(name);
        }

        public static CellMap GetCellMapForShapes()
        {
            if (shape_cellmap == null)
            {
                shape_cellmap = new CellMap();
                shape_cellmap["Angle"] = VA.ShapeSheet.SRCConstants.Angle;
                shape_cellmap["BeginX"] = VA.ShapeSheet.SRCConstants.BeginX;
                shape_cellmap["BeginY"] = VA.ShapeSheet.SRCConstants.BeginY;
                shape_cellmap["CharCase"] = VA.ShapeSheet.SRCConstants.CharCase;
                shape_cellmap["CharColor"] = VA.ShapeSheet.SRCConstants.CharColor;
                shape_cellmap["CharColorTransparency"] = VA.ShapeSheet.SRCConstants.CharColorTrans;
                shape_cellmap["CharFont"] = VA.ShapeSheet.SRCConstants.CharFont;
                shape_cellmap["CharFontScale"] = VA.ShapeSheet.SRCConstants.CharFontScale;
                shape_cellmap["CharLetterspace"] = VA.ShapeSheet.SRCConstants.CharLetterspace;
                shape_cellmap["CharSize"] = VA.ShapeSheet.SRCConstants.CharSize;
                shape_cellmap["CharStyle"] = VA.ShapeSheet.SRCConstants.CharStyle;
                shape_cellmap["EndX"] = VA.ShapeSheet.SRCConstants.EndX;
                shape_cellmap["EndY"] = VA.ShapeSheet.SRCConstants.EndY;
                shape_cellmap["FillBkgnd"] = VA.ShapeSheet.SRCConstants.FillBkgnd;
                shape_cellmap["FillBkgndTrans"] = VA.ShapeSheet.SRCConstants.FillBkgndTrans;
                shape_cellmap["FillForegnd"] = VA.ShapeSheet.SRCConstants.FillForegnd;
                shape_cellmap["FillForegndTrans"] = VA.ShapeSheet.SRCConstants.FillForegndTrans;
                shape_cellmap["FillPattern"] = VA.ShapeSheet.SRCConstants.FillPattern;
                shape_cellmap["Height"] = VA.ShapeSheet.SRCConstants.Height;
                shape_cellmap["LineCap"] = VA.ShapeSheet.SRCConstants.LineCap;
                shape_cellmap["LineColor"] = VA.ShapeSheet.SRCConstants.LineColor;
                shape_cellmap["LinePattern"] = VA.ShapeSheet.SRCConstants.LinePattern;
                shape_cellmap["LineWeight"] = VA.ShapeSheet.SRCConstants.LineWeight;
                shape_cellmap["LockAspect"] = VA.ShapeSheet.SRCConstants.LockAspect;
                shape_cellmap["LockBegin"] = VA.ShapeSheet.SRCConstants.LockBegin;
                shape_cellmap["LockCalcWH"] = VA.ShapeSheet.SRCConstants.LockCalcWH;
                shape_cellmap["LockCrop"] = VA.ShapeSheet.SRCConstants.LockCrop;
                shape_cellmap["LockCustProp"] = VA.ShapeSheet.SRCConstants.LockCustProp;
                shape_cellmap["LockDelete"] = VA.ShapeSheet.SRCConstants.LockDelete;
                shape_cellmap["LockEnd"] = VA.ShapeSheet.SRCConstants.LockEnd;
                shape_cellmap["LockFormat"] = VA.ShapeSheet.SRCConstants.LockFormat;
                shape_cellmap["LockFromGroupFormat"] = VA.ShapeSheet.SRCConstants.LockFromGroupFormat;
                shape_cellmap["LockGroup"] = VA.ShapeSheet.SRCConstants.LockGroup;
                shape_cellmap["LockHeight"] = VA.ShapeSheet.SRCConstants.LockHeight;
                shape_cellmap["LockMoveX"] = VA.ShapeSheet.SRCConstants.LockMoveX;
                shape_cellmap["LockMoveY"] = VA.ShapeSheet.SRCConstants.LockMoveY;
                shape_cellmap["LockRotate"] = VA.ShapeSheet.SRCConstants.LockRotate;
                shape_cellmap["LockSelect"] = VA.ShapeSheet.SRCConstants.LockSelect;
                shape_cellmap["LockTextEdit"] = VA.ShapeSheet.SRCConstants.LockTextEdit;
                shape_cellmap["LockThemeColors"] = VA.ShapeSheet.SRCConstants.LockThemeColors;
                shape_cellmap["LockThemeEffects"] = VA.ShapeSheet.SRCConstants.LockThemeEffects;
                shape_cellmap["LockVtxEdit"] = VA.ShapeSheet.SRCConstants.LockVtxEdit;
                shape_cellmap["LockWidth"] = VA.ShapeSheet.SRCConstants.LockWidth;
                shape_cellmap["LocPinX"] = VA.ShapeSheet.SRCConstants.LocPinX;
                shape_cellmap["LocPinY"] = VA.ShapeSheet.SRCConstants.LocPinY;
                shape_cellmap["PinX"] = VA.ShapeSheet.SRCConstants.PinX;
                shape_cellmap["PinY"] = VA.ShapeSheet.SRCConstants.PinY;
                shape_cellmap["Rounding"] = VA.ShapeSheet.SRCConstants.Rounding;
                shape_cellmap["SelectMode"] = VA.ShapeSheet.SRCConstants.SelectMode;
                shape_cellmap["ShdwBkgnd"] = VA.ShapeSheet.SRCConstants.ShdwBkgnd;
                shape_cellmap["ShdwBkgndTrans"] = VA.ShapeSheet.SRCConstants.ShdwBkgndTrans;
                shape_cellmap["ShdwForegnd"] = VA.ShapeSheet.SRCConstants.ShdwForegnd;
                shape_cellmap["ShdwForegndTrans"] = VA.ShapeSheet.SRCConstants.ShdwForegndTrans;
                shape_cellmap["ShdwObliqueAngle"] = VA.ShapeSheet.SRCConstants.ShdwObliqueAngle;
                shape_cellmap["ShdwOffsetX"] = VA.ShapeSheet.SRCConstants.ShdwOffsetX;
                shape_cellmap["ShdwOffsetY"] = VA.ShapeSheet.SRCConstants.ShdwOffsetY;
                shape_cellmap["ShdwPattern"] = VA.ShapeSheet.SRCConstants.ShdwPattern;
                shape_cellmap["ShdwScaleFactor"] = VA.ShapeSheet.SRCConstants.ShdwScaleFactor;
                shape_cellmap["ShdwType"] = VA.ShapeSheet.SRCConstants.ShdwType;
                shape_cellmap["TxtAngle"] = VA.ShapeSheet.SRCConstants.TxtAngle;
                shape_cellmap["TxtHeight"] = VA.ShapeSheet.SRCConstants.TxtHeight;
                shape_cellmap["TxtLocPinX"] = VA.ShapeSheet.SRCConstants.TxtLocPinX;
                shape_cellmap["TxtLocPinY"] = VA.ShapeSheet.SRCConstants.TxtLocPinY;
                shape_cellmap["TxtPinX"] = VA.ShapeSheet.SRCConstants.TxtPinX;
                shape_cellmap["TxtPinY"] = VA.ShapeSheet.SRCConstants.TxtPinY;
                shape_cellmap["TxtWidth"] = VA.ShapeSheet.SRCConstants.TxtWidth;
                shape_cellmap["Width"] = VA.ShapeSheet.SRCConstants.Width;

            }
            return shape_cellmap;
        }

        public static CellMap GetCellMapForPages()
        {
            if (page_cellmap == null)
            {
                page_cellmap = new CellMap();
                page_cellmap["PageBottomMargin"] = VA.ShapeSheet.SRCConstants.PageBottomMargin;
                page_cellmap["PageHeight"] = VA.ShapeSheet.SRCConstants.PageHeight;
                page_cellmap["PageLeftMargin"] = VA.ShapeSheet.SRCConstants.PageLeftMargin;
                page_cellmap["PageLineJumpDirX"] = VA.ShapeSheet.SRCConstants.PageLineJumpDirX;
                page_cellmap["PageLineJumpDirY"] = VA.ShapeSheet.SRCConstants.PageLineJumpDirY;

                page_cellmap["PageRightMargin"] = VA.ShapeSheet.SRCConstants.PageRightMargin;
                page_cellmap["PageScale"] = VA.ShapeSheet.SRCConstants.PageScale;
                page_cellmap["PageShapeSplit"] = VA.ShapeSheet.SRCConstants.PageShapeSplit;
                page_cellmap["PageTopMargin"] = VA.ShapeSheet.SRCConstants.PageTopMargin;
                page_cellmap["PageWidth"] = VA.ShapeSheet.SRCConstants.PageWidth;
                page_cellmap["CenterX"] = VA.ShapeSheet.SRCConstants.CenterX;
                page_cellmap["CenterY"] = VA.ShapeSheet.SRCConstants.CenterY;
                page_cellmap["PaperKind"] = VA.ShapeSheet.SRCConstants.PaperKind;
                page_cellmap["PrintGrid"] = VA.ShapeSheet.SRCConstants.PrintGrid;
                page_cellmap["PrintPageOrientation"] = VA.ShapeSheet.SRCConstants.PrintPageOrientation;
                page_cellmap["ScaleX"] = VA.ShapeSheet.SRCConstants.ScaleX;
                page_cellmap["ScaleY"] = VA.ShapeSheet.SRCConstants.ScaleY;
                page_cellmap["PaperSource"] = VA.ShapeSheet.SRCConstants.PaperSource;
                page_cellmap["DrawingScale"] = VA.ShapeSheet.SRCConstants.DrawingScale;
                page_cellmap["DrawingScaleType"] = VA.ShapeSheet.SRCConstants.DrawingScaleType;
                page_cellmap["DrawingSizeType"] = VA.ShapeSheet.SRCConstants.DrawingSizeType;
                page_cellmap["InhibitSnap"] = VA.ShapeSheet.SRCConstants.InhibitSnap;
                page_cellmap["ShdwObliqueAngle"] = VA.ShapeSheet.SRCConstants.ShdwObliqueAngle;
                page_cellmap["ShdwOffsetX"] = VA.ShapeSheet.SRCConstants.ShdwOffsetX;
                page_cellmap["ShdwOffsetY"] = VA.ShapeSheet.SRCConstants.ShdwOffsetY;
                page_cellmap["ShdwScaleFactor"] = VA.ShapeSheet.SRCConstants.ShdwScaleFactor;
                page_cellmap["ShdwType"] = VA.ShapeSheet.SRCConstants.ShdwType;
                page_cellmap["UIVisibility"] = VA.ShapeSheet.SRCConstants.UIVisibility;
                page_cellmap["XGridDensity"] = VA.ShapeSheet.SRCConstants.XGridDensity;
                page_cellmap["XGridOrigin"] = VA.ShapeSheet.SRCConstants.XGridOrigin;
                page_cellmap["XGridSpacing"] = VA.ShapeSheet.SRCConstants.XGridSpacing;
                page_cellmap["XRulerDensity"] = VA.ShapeSheet.SRCConstants.XRulerDensity;
                page_cellmap["XRulerOrigin"] = VA.ShapeSheet.SRCConstants.XRulerOrigin;
                page_cellmap["YGridDensity"] = VA.ShapeSheet.SRCConstants.YGridDensity;
                page_cellmap["YGridOrigin"] = VA.ShapeSheet.SRCConstants.YGridOrigin;
                page_cellmap["YGridSpacing"] = VA.ShapeSheet.SRCConstants.YGridSpacing;
                page_cellmap["YRulerDensity"] = VA.ShapeSheet.SRCConstants.YRulerDensity;
                page_cellmap["YRulerOrigin"] = VA.ShapeSheet.SRCConstants.YRulerOrigin;
                page_cellmap["AvenueSizeX"] = VA.ShapeSheet.SRCConstants.AvenueSizeX;
                page_cellmap["AvenueSizeY"] = VA.ShapeSheet.SRCConstants.AvenueSizeY;
                page_cellmap["BlockSizeX"] = VA.ShapeSheet.SRCConstants.BlockSizeX;
                page_cellmap["BlockSizeY"] = VA.ShapeSheet.SRCConstants.BlockSizeY;
                page_cellmap["CtrlAsInput"] = VA.ShapeSheet.SRCConstants.CtrlAsInput;
                page_cellmap["DynamicsOff"] = VA.ShapeSheet.SRCConstants.DynamicsOff;
                page_cellmap["EnableGrid"] = VA.ShapeSheet.SRCConstants.EnableGrid;
                page_cellmap["LineAdjustFrom"] = VA.ShapeSheet.SRCConstants.LineAdjustFrom;
                page_cellmap["LineAdjustTo"] = VA.ShapeSheet.SRCConstants.LineAdjustTo;
                page_cellmap["LineJumpCode"] = VA.ShapeSheet.SRCConstants.LineJumpCode;
                page_cellmap["LineJumpFactorX"] = VA.ShapeSheet.SRCConstants.LineJumpFactorX;
                page_cellmap["LineJumpFactorY"] = VA.ShapeSheet.SRCConstants.LineJumpFactorY;
                page_cellmap["LineJumpStyle"] = VA.ShapeSheet.SRCConstants.LineJumpStyle;
                page_cellmap["LineRouteExt"] = VA.ShapeSheet.SRCConstants.LineRouteExt;
                page_cellmap["LineToLineX"] = VA.ShapeSheet.SRCConstants.LineToLineX;
                page_cellmap["LineToLineY"] = VA.ShapeSheet.SRCConstants.LineToLineY;
                page_cellmap["LineToNodeX"] = VA.ShapeSheet.SRCConstants.LineToNodeX;
                page_cellmap["LineToNodeY"] = VA.ShapeSheet.SRCConstants.LineToNodeY;
                page_cellmap["PlaceDepth"] = VA.ShapeSheet.SRCConstants.PlaceDepth;
                page_cellmap["PlaceFlip"] = VA.ShapeSheet.SRCConstants.PlaceFlip;
                page_cellmap["PlaceStyle"] = VA.ShapeSheet.SRCConstants.PlaceStyle;
                page_cellmap["PlowCode"] = VA.ShapeSheet.SRCConstants.PlowCode;
                page_cellmap["ResizePage"] = VA.ShapeSheet.SRCConstants.ResizePage;
                page_cellmap["RouteStyle"] = VA.ShapeSheet.SRCConstants.RouteStyle;
                page_cellmap["AvoidPageBreaks"] = VA.ShapeSheet.SRCConstants.AvoidPageBreaks;
                page_cellmap["DrawingResizeType"] = VA.ShapeSheet.SRCConstants.DrawingResizeType;

            }
            return page_cellmap;
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
                    query.AddCell(name_to_src[resolved_cellname], resolved_cellname);
                }
            }
            return query;
        }
    }
}