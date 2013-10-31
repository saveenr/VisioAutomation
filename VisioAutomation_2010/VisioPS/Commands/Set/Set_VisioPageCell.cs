using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using System.Linq;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioPageCell")]
    public class Set_VisioPageCell: VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false,Position=0)] 
        public System.Collections.Hashtable Hashtable  { get; set; }

        [SMA.Parameter(Mandatory = false)] 
        public string PageWidth { get; set; }
        
        [SMA.Parameter(Mandatory = false)] 
        public string PageHeight { get; set; }
        
        [SMA.Parameter(Mandatory = false)] 
        public string PageBottomMargin;
        
        [SMA.Parameter(Mandatory = false)]
        public string PageLeftMargin { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PageLineJumpDirX { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PageLineJumpDirY { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PageRightMargin { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PageScale { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PageShapeSplit { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PageTopMargin { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string CenterX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string CenterY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string PaperKind { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PrintGrid { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PrintPageOrientation { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string ScaleX { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string ScaleY { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PaperSource { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string DrawingScaleType { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string DrawingScale { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string DrawingSizeType { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string InhibitSnap { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string ShdwObliqueAngle { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string ShdwOffsetX { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string ShdwOffsetY { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string ShdwScaleFactor { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string ShdwType { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string UIVisibility { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string XGridDensity { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string XGridOrigin { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string XGridSpacing { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string XRulerDensity { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string XRulerOrigin { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string YGridDensity { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string YGridOrigin { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string YGridSpacing { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string YRulerDensity { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string YRulerOrigin { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string AvenueSizeX { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string AvenueSizeY { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string BlockSizeX { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string BlockSizeY { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string CtrlAsInput { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string DynamicsOff { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string EnableGrid { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineAdjustFrom { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineAdjustTo { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineJumpCode { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineJumpFactorX { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineJumpFactorY { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineJumpStyle { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineRouteExt { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineToLineX { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineToLineY { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineToNodeX { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string LineToNodeY { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PlaceDepth { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PlaceFlip { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PlaceStyle { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string ResizePage { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string PlowCode { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string RouteStyle { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string AvoidPageBreaks { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public string DrawingResizeType { get; set; }
 
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page[] Pages { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular { get; set; }
        
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            var update = new VisioAutomation.ShapeSheet.Update();
            update.BlastGuards = this.BlastGuards;
            update.TestCircular= this.TestCircular;

            var target_pages = this.Pages ?? new [] { scriptingsession.Page.Get() };

            var valuemap = new CellValueMap(Get_VisioPageCell.GetPageCellDictionary());

            valuemap.UpdateValueMap(this.Hashtable);

            valuemap.SetIf("PageBottomMargin",this.PageBottomMargin);
            valuemap.SetIf("PageHeight",this.PageHeight);
            valuemap.SetIf("PageLeftMargin",this.PageLeftMargin);
            valuemap.SetIf("PageLineJumpDirX",this.PageLineJumpDirX);
            valuemap.SetIf("PageLineJumpDirY",this.PageLineJumpDirY);
            valuemap.SetIf("PageRightMargin",this.PageRightMargin);
            valuemap.SetIf("PageScale",this.PageScale);
            valuemap.SetIf("PageShapeSplit",this.PageShapeSplit);
            valuemap.SetIf("PageTopMargin",this.PageTopMargin);
            valuemap.SetIf("PageWidth",this.PageWidth);
            valuemap.SetIf("CenterX",this.CenterX);
            valuemap.SetIf("CenterY",this.CenterY);
            valuemap.SetIf("PaperKind",this.PaperKind);
            valuemap.SetIf("PrintGrid",this.PrintGrid);
            valuemap.SetIf("PrintPageOrientation",this.PrintPageOrientation);
            valuemap.SetIf("ScaleX",this.ScaleX);
            valuemap.SetIf("ScaleY",this.ScaleY);
            valuemap.SetIf("PaperSource",this.PaperSource);
            valuemap.SetIf("DrawingScale",this.DrawingScale);
            valuemap.SetIf("DrawingScaleType",this.DrawingScaleType);
            valuemap.SetIf("DrawingSizeType",this.DrawingSizeType);
            valuemap.SetIf("InhibitSnap",this.InhibitSnap);
            valuemap.SetIf("ShdwObliqueAngle",this.ShdwObliqueAngle);
            valuemap.SetIf("ShdwOffsetX",this.ShdwOffsetX);
            valuemap.SetIf("ShdwOffsetY",this.ShdwOffsetY);
            valuemap.SetIf("ShdwScaleFactor",this.ShdwScaleFactor);
            valuemap.SetIf("ShdwType",this.ShdwType);
            valuemap.SetIf("UIVisibility",this.UIVisibility);
            valuemap.SetIf("XGridDensity",this.XGridDensity);
            valuemap.SetIf("XGridOrigin",this.XGridOrigin);
            valuemap.SetIf("XGridSpacing",this.XGridSpacing);
            valuemap.SetIf("XRulerDensity",this.XRulerDensity);
            valuemap.SetIf("XRulerOrigin",this.XRulerOrigin);
            valuemap.SetIf("YGridDensity",this.YGridDensity);
            valuemap.SetIf("YGridOrigin",this.YGridOrigin);
            valuemap.SetIf("YGridSpacing",this.YGridSpacing);
            valuemap.SetIf("YRulerDensity",this.YRulerDensity);
            valuemap.SetIf("YRulerOrigin",this.YRulerOrigin);
            valuemap.SetIf("AvenueSizeX",this.AvenueSizeX);
            valuemap.SetIf("AvenueSizeY",this.AvenueSizeY);
            valuemap.SetIf("BlockSizeX",this.BlockSizeX);
            valuemap.SetIf("BlockSizeY",this.BlockSizeY);
            valuemap.SetIf("CtrlAsInput",this.CtrlAsInput);
            valuemap.SetIf("DynamicsOff",this.DynamicsOff);
            valuemap.SetIf("EnableGrid",this.EnableGrid);
            valuemap.SetIf("LineAdjustFrom",this.LineAdjustFrom);
            valuemap.SetIf("LineAdjustTo",this.LineAdjustTo);
            valuemap.SetIf("LineJumpCode",this.LineJumpCode);
            valuemap.SetIf("LineJumpFactorX",this.LineJumpFactorX);
            valuemap.SetIf("LineJumpFactorY",this.LineJumpFactorY);
            valuemap.SetIf("LineJumpStyle",this.LineJumpStyle);
            valuemap.SetIf("LineRouteExt",this.LineRouteExt);
            valuemap.SetIf("LineToLineX",this.LineToLineX);
            valuemap.SetIf("LineToLineY",this.LineToLineY);
            valuemap.SetIf("LineToNodeX",this.LineToNodeX);
            valuemap.SetIf("LineToNodeY",this.LineToNodeY);
            valuemap.SetIf("PageLineJumpDirX",this.PageLineJumpDirX);
            valuemap.SetIf("PageLineJumpDirY",this.PageLineJumpDirY);
            valuemap.SetIf("PageShapeSplit",this.PageShapeSplit);
            valuemap.SetIf("PlaceDepth",this.PlaceDepth);
            valuemap.SetIf("PlaceFlip",this.PlaceFlip);
            valuemap.SetIf("PlaceStyle",this.PlaceStyle);
            valuemap.SetIf("PlowCode",this.PlowCode);
            valuemap.SetIf("ResizePage",this.ResizePage);
            valuemap.SetIf("RouteStyle",this.RouteStyle);
            valuemap.SetIf("AvoidPageBreaks",this.AvoidPageBreaks);
            valuemap.SetIf("DrawingResizeType",this.DrawingResizeType);


            foreach (var page in target_pages)
            {
                var pagesheet = page.PageSheet;

                foreach (var cellname in valuemap.CellNames)
                {
                    string cell_value = valuemap[cellname];
                    var cell_src = valuemap.GetSRC(cellname);
                    update.SetFormulaIgnoreNull( cell_src , cell_value);
                }
                this.WriteVerboseEx("BlastGuards: {0}", this.BlastGuards);
                this.WriteVerboseEx("TestCircular: {0}", this.TestCircular);
                this.WriteVerboseEx("Number of Shapes : {0}", 1);
                this.WriteVerboseEx("Number of Total Updates: {0}", update.Count());
                this.WriteVerboseEx("Number of Updates per Shape: {0}", update.Count() / 1);

                using (var undoscope = new VA.Application.UndoScope(this.ScriptingSession.VisioApplication, "SetPageCells"))
                {
                    this.WriteVerboseEx("Start Update");
                    update.Execute(pagesheet);
                    this.WriteVerboseEx("End Update");
                }
            }

        }

    }

    public class CellValueMap
    {
        Dictionary<string, string> dic;

        private System.Text.RegularExpressions.Regex regex_cellname;
        private System.Text.RegularExpressions.Regex regex_cellname_wildcard;

        private CellMap srcmap;

        public CellValueMap( CellMap srcMap)
        {
            this.regex_cellname = new System.Text.RegularExpressions.Regex("^[a-zA-Z]*$");
            this.regex_cellname_wildcard = new System.Text.RegularExpressions.Regex("^[a-zA-Z\\*\\?]*$");
            this.dic = new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase);
            this.srcmap = srcMap;
        }

        public VA.ShapeSheet.SRC GetSRC(string name)
        {
            return this.srcmap[name];
        }

        public string this[string name]
        {
            get { return this.dic[name]; }
            set
            {
                this.CheckCellName(name);
                this.dic[name] = value;
            }
        }

        public void SetIf(string name, string value)
        {
            if (value != null)
            {
                this.dic[name] = value;
            }            
        }

        public void SetIf(int id, string name, string value)
        {
            if (value != null)
            {
                this.dic[name] = value;
            }

        }

        public Dictionary<string, string>.KeyCollection CellNames
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
        }


        public void UpdateValueMap(System.Collections.Hashtable Hashtable)
        {
            if (Hashtable != null)
            {
                foreach (object key_o in Hashtable.Keys)
                {
                    if (!(key_o is string))
                    {
                        string message = "Only string values can be keys in the hashtable";
                        throw new System.ArgumentOutOfRangeException(message);
                    }
                    string key_string = (string)key_o;

                    object value_o = Hashtable[key_o];
                    if (value_o == null)
                    {
                        string message = "Null values not allowed for cellvalues";
                        throw new System.ArgumentOutOfRangeException(message);
                    }
                    if (value_o is string)
                    {
                        string value_string = (string)value_o;
                        this[key_string] = value_string;
                    }
                    else if (value_o is int)
                    {
                        int value_int = (int)value_o;
                        string value_string = value_int.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        this[key_string] = value_string;
                    }
                    else if (value_o is float)
                    {
                        float value_float = (float)value_o;
                        string value_string = value_float.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        this[key_string] = value_string;
                    }
                    else if (value_o is double)
                    {
                        double value_double = (double)value_o;
                        string value_string = value_double.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        this[key_string] = value_string;
                    }
                    else
                    {
                        string message = string.Format("Cell values cannot be of type {0} ", value_o.GetType().Name);
                        throw new System.ArgumentOutOfRangeException(message);
                    }
                }
            }
        }

    }

}