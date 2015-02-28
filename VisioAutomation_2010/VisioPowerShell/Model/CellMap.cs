using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;

namespace VisioPowerShell
{
    public class CellMap
    {
        Dictionary<string, VA.ShapeSheet.SRC> dic;

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
            get
            {
                return this.dic.Keys;
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

            string msg = string.Format("Cell name \"{0}\" is not valid",name);
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

                var regex = new System.Text.RegularExpressions.Regex(pat, System.Text.RegularExpressions.RegexOptions.IgnoreCase);

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

        private static CellMap map_name_to_shape_cell;
        private static CellMap map_name_to_page_cell;


        public static CellMap GetShapeCellDictionary()
        {
            if (map_name_to_shape_cell == null)
            {
                map_name_to_shape_cell = new CellMap();
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.Angle.Name] = VA.ShapeSheet.SRCConstants.Angle;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.BeginX.Name] = VA.ShapeSheet.SRCConstants.BeginX;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.BeginY.Name] = VA.ShapeSheet.SRCConstants.BeginY;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.CharCase.Name] = VA.ShapeSheet.SRCConstants.CharCase;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.CharColor.Name] = VA.ShapeSheet.SRCConstants.CharColor;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.CharColorTrans.Name] = VA.ShapeSheet.SRCConstants.CharColorTrans;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.CharFont.Name] = VA.ShapeSheet.SRCConstants.CharFont;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.CharFontScale.Name] = VA.ShapeSheet.SRCConstants.CharFontScale;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.CharLetterspace.Name] = VA.ShapeSheet.SRCConstants.CharLetterspace;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.CharSize.Name] = VA.ShapeSheet.SRCConstants.CharSize;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.CharStyle.Name] = VA.ShapeSheet.SRCConstants.CharStyle;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.EndX.Name] = VA.ShapeSheet.SRCConstants.EndX;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.EndY.Name] = VA.ShapeSheet.SRCConstants.EndY;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.FillBkgnd.Name] = VA.ShapeSheet.SRCConstants.FillBkgnd;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.FillBkgndTrans.Name] = VA.ShapeSheet.SRCConstants.FillBkgndTrans;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.FillForegnd.Name] = VA.ShapeSheet.SRCConstants.FillForegnd;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.FillForegndTrans.Name] = VA.ShapeSheet.SRCConstants.FillForegndTrans;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.FillPattern.Name] = VA.ShapeSheet.SRCConstants.FillPattern;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.Height.Name] = VA.ShapeSheet.SRCConstants.Height;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LineCap.Name] = VA.ShapeSheet.SRCConstants.LineCap;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LineColor.Name] = VA.ShapeSheet.SRCConstants.LineColor;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LinePattern.Name] = VA.ShapeSheet.SRCConstants.LinePattern;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LineWeight.Name] = VA.ShapeSheet.SRCConstants.LineWeight;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockAspect.Name] = VA.ShapeSheet.SRCConstants.LockAspect;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockBegin.Name] = VA.ShapeSheet.SRCConstants.LockBegin;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockCalcWH.Name] = VA.ShapeSheet.SRCConstants.LockCalcWH;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockCrop.Name] = VA.ShapeSheet.SRCConstants.LockCrop;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockCustProp.Name] = VA.ShapeSheet.SRCConstants.LockCustProp;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockDelete.Name] = VA.ShapeSheet.SRCConstants.LockDelete;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockEnd.Name] = VA.ShapeSheet.SRCConstants.LockEnd;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockFormat.Name] = VA.ShapeSheet.SRCConstants.LockFormat;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockFromGroupFormat.Name] = VA.ShapeSheet.SRCConstants.LockFromGroupFormat;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockGroup.Name] = VA.ShapeSheet.SRCConstants.LockGroup;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockHeight.Name] = VA.ShapeSheet.SRCConstants.LockHeight;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockMoveX.Name] = VA.ShapeSheet.SRCConstants.LockMoveX;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockMoveY.Name] = VA.ShapeSheet.SRCConstants.LockMoveY;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockRotate.Name] = VA.ShapeSheet.SRCConstants.LockRotate;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockSelect.Name] = VA.ShapeSheet.SRCConstants.LockSelect;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockTextEdit.Name] = VA.ShapeSheet.SRCConstants.LockTextEdit;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockThemeColors.Name] = VA.ShapeSheet.SRCConstants.LockThemeColors;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockThemeEffects.Name] = VA.ShapeSheet.SRCConstants.LockThemeEffects;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockVtxEdit.Name] = VA.ShapeSheet.SRCConstants.LockVtxEdit;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LockWidth.Name] = VA.ShapeSheet.SRCConstants.LockWidth;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LocPinX.Name] = VA.ShapeSheet.SRCConstants.LocPinX;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.LocPinY.Name] = VA.ShapeSheet.SRCConstants.LocPinY;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.PinX.Name] = VA.ShapeSheet.SRCConstants.PinX;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.PinY.Name] = VA.ShapeSheet.SRCConstants.PinY;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.Rounding.Name] = VA.ShapeSheet.SRCConstants.Rounding;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.SelectMode.Name] = VA.ShapeSheet.SRCConstants.SelectMode;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.ShdwBkgnd.Name] = VA.ShapeSheet.SRCConstants.ShdwBkgnd;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.ShdwBkgndTrans.Name] = VA.ShapeSheet.SRCConstants.ShdwBkgndTrans;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.ShdwForegnd.Name] = VA.ShapeSheet.SRCConstants.ShdwForegnd;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.ShdwForegndTrans.Name] = VA.ShapeSheet.SRCConstants.ShdwForegndTrans;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.ShdwObliqueAngle.Name] = VA.ShapeSheet.SRCConstants.ShdwObliqueAngle;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.ShdwOffsetX.Name] = VA.ShapeSheet.SRCConstants.ShdwOffsetX;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.ShdwOffsetY.Name] = VA.ShapeSheet.SRCConstants.ShdwOffsetY;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.ShdwPattern.Name] = VA.ShapeSheet.SRCConstants.ShdwPattern;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.ShdwScaleFactor.Name] = VA.ShapeSheet.SRCConstants.ShdwScaleFactor;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.ShdwType.Name] = VA.ShapeSheet.SRCConstants.ShdwType;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.TxtAngle.Name] = VA.ShapeSheet.SRCConstants.TxtAngle;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.TxtHeight.Name] = VA.ShapeSheet.SRCConstants.TxtHeight;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.TxtLocPinX.Name] = VA.ShapeSheet.SRCConstants.TxtLocPinX;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.TxtLocPinY.Name] = VA.ShapeSheet.SRCConstants.TxtLocPinY;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.TxtPinX.Name] = VA.ShapeSheet.SRCConstants.TxtPinX;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.TxtPinY.Name] = VA.ShapeSheet.SRCConstants.TxtPinY;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.TxtWidth.Name] = VA.ShapeSheet.SRCConstants.TxtWidth;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.Width.Name] = VA.ShapeSheet.SRCConstants.Width;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.BeginArrow.Name] = VA.ShapeSheet.SRCConstants.BeginArrow;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.BeginArrowSize.Name] = VA.ShapeSheet.SRCConstants.BeginArrowSize;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.EndArrow.Name] = VA.ShapeSheet.SRCConstants.EndArrow;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.EndArrowSize.Name] = VA.ShapeSheet.SRCConstants.EndArrowSize;
                map_name_to_shape_cell[VA.ShapeSheet.SRCConstants.HideText.Name] = VA.ShapeSheet.SRCConstants.HideText;
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


    }
}