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
                var list = GetShapeCells();
                foreach (var src in list)
                {
                    map_name_to_shape_cell[src.Name] = src;
                }
            }
            return map_name_to_shape_cell;
        }

        public static CellMap GetPageCellDictionary()
        {
            if (map_name_to_page_cell == null)
            {
                map_name_to_page_cell = new CellMap();
                var list = GetPageCells();
                foreach (var src in list)
                {
                    map_name_to_shape_cell[src.Name] = src;
                }
            }
            return map_name_to_page_cell;
        }
    }
}