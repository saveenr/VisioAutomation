using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using System.Linq;
using VA = VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioShapeCell")]
    public class Get_VisioShapeCell : VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public string[] Cells { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Width { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Height { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PinX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter PinY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LocPinX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LocPinY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Angle { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter FillPattern { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter FillForegnd { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter FillForegndTrans { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter FillBkgnd { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter FillBkgndTrans { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LinePattern { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LineWeight { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LineColor { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LineCap { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Rounding { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter CharCase { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter CharColor { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter CharFont { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter CharFontScale { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter CharLetterspace { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter CharSize { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter CharStyle { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter CharColorTransparency { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BeginArrow { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BeginArrowSize { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter EndArrow { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter EndArrowSize { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BeginX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BeginY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter EndX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter EndY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ShdwBkgnd { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ShdwBkgndTrans { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ShdwForegnd { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ShdwForegndTrans { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ShdwObliqueAngle { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ShdwOffsetX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ShdwOffsetY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ShdwPattern { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ShdwScalefactor { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ShdwType { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter SelectMode { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockAspect { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockBegin { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockCalcWH { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockCrop { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockCustProp { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockDelete { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockEnd { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockFormat { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockFromGroupFormat { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockGroup { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockHeight { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockMoveX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockMoveY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockRotate { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockSelect { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockTextEdit { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockThemeColors { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockThemeEffects { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockVtxEdit { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter LockWidth { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TxtAngle { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TxtHeight { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TxtLocPinX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TxtLocPinY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TxtPinX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TxtPinY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TxtWidth { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter HideText { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes { get; set; }

        [SMA.Parameter(Mandatory = false)] 
        public SMA.SwitchParameter GetResults;

        [SMA.Parameter(Mandatory = false)] 
        public ResultType ResultType = ResultType.String;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            var query = new VisioAutomation.ShapeSheet.Query.CellQuery();

            var target_shapes = this.Shapes ?? scriptingsession.Selection.GetShapes();
            var target_shapeids = target_shapes.Select(s => s.ID).ToList();

            addcell(query, this.Angle, "Angle");
            addcell(query, this.BeginArrow, "BeginArrow");
            addcell(query, this.BeginArrowSize, "BeginArrowSize");
            addcell(query, this.BeginX, "BeginX");
            addcell(query, this.BeginY, "BeginY");
            addcell(query, this.CharCase, "CharCase");
            addcell(query, this.CharColor, "CharColor");
            addcell(query, this.CharColorTransparency, "CharColorTransparency");
            addcell(query, this.CharFont, "CharFont");
            addcell(query, this.CharFontScale, "CharFontScale");
            addcell(query, this.CharLetterspace, "CharLetterspace");
            addcell(query, this.CharSize, "CharSize");
            addcell(query, this.CharStyle, "CharStyle");
            addcell(query, this.EndArrow, "EndArrow");
            addcell(query, this.EndArrowSize, "EndArrowSize");
            addcell(query, this.EndX, "EndX");
            addcell(query, this.EndY, "EndY");
            addcell(query, this.FillBkgnd, "FillBkgnd");
            addcell(query, this.FillBkgndTrans, "FillBkgndTrans");
            addcell(query, this.FillForegnd, "FillForegnd");
            addcell(query, this.FillForegndTrans, "FillForegndTrans");
            addcell(query, this.FillPattern, "FillPattern");
            addcell(query, this.Height, "Height");
            addcell(query, this.HideText, "HideText");
            addcell(query, this.LineCap, "LineCap");
            addcell(query, this.LineColor, "LineColor");
            addcell(query, this.LinePattern, "LinePattern");
            addcell(query, this.LineWeight, "LineWeight");
            addcell(query, this.LockAspect, "LockAspect");
            addcell(query, this.LockBegin, "LockBegin");
            addcell(query, this.LockCalcWH, "LockCalcWH");
            addcell(query, this.LockCrop, "LockCrop");
            addcell(query, this.LockCustProp, "LockCustProp");
            addcell(query, this.LockDelete, "LockDelete");
            addcell(query, this.LockEnd, "LockEnd");
            addcell(query, this.LockFormat, "LockFormat");
            addcell(query, this.LockFromGroupFormat, "LockFromGroupFormat");
            addcell(query, this.LockGroup, "LockGroup");
            addcell(query, this.LockHeight, "LockHeight");
            addcell(query, this.LockMoveX, "LockMoveX");
            addcell(query, this.LockMoveY, "LockMoveY");
            addcell(query, this.LockRotate, "LockRotate");
            addcell(query, this.LockSelect, "LockSelect");
            addcell(query, this.LockTextEdit, "LockTextEdit");
            addcell(query, this.LockThemeColors, "LockThemeColors");
            addcell(query, this.LockThemeEffects, "LockThemeEffects");
            addcell(query, this.LockVtxEdit, "LockVtxEdit");
            addcell(query, this.LockWidth, "LockWidth");
            addcell(query, this.LocPinX, "LocPinX");
            addcell(query, this.LocPinY, "LocPinY");
            addcell(query, this.PinX, "PinX");
            addcell(query, this.PinY, "PinY");
            addcell(query, this.Rounding, "Rounding");
            addcell(query, this.SelectMode, "SelectMode");
            addcell(query, this.ShdwBkgnd, "ShdwBkgnd");
            addcell(query, this.ShdwBkgndTrans, "ShdwBkgndTrans");
            addcell(query, this.ShdwForegnd, "ShdwForegnd");
            addcell(query, this.ShdwForegndTrans, "ShdwForegndTrans");
            addcell(query, this.ShdwObliqueAngle, "ShdwObliqueAngle");
            addcell(query, this.ShdwOffsetX, "ShdwOffsetX");
            addcell(query, this.ShdwOffsetY, "ShdwOffsetY");
            addcell(query, this.ShdwPattern, "ShdwPattern");
            addcell(query, this.ShdwScalefactor, "ShdwScalefactor");
            addcell(query, this.ShdwType, "ShdwType");
            addcell(query, this.TxtAngle, "TxtAngle");
            addcell(query, this.TxtHeight, "TxtHeight");
            addcell(query, this.TxtLocPinX, "TxtLocPinX");
            addcell(query, this.TxtLocPinY, "TxtLocPinY");
            addcell(query, this.TxtPinX, "TxtPinX");
            addcell(query, this.TxtPinY, "TxtPinY");
            addcell(query, this.TxtWidth, "TxtWidth");
            addcell(query, this.Width, "Width");

            var dic = this.GetShapeCellDictionary();
            Get_VisioPageCell.SetFromCellNames(query, this.Cells, dic);

            var page = scriptingsession.Page.Get();

            this.WriteVerboseEx("Number of Shapes : {0}", target_shapes.Count);
            this.WriteVerboseEx("Number of Cells: {0}", query.Columns.Count);

            this.WriteVerboseEx("Start Query");

            var dt = VisioPSUtil.QueryToDataTable(query, this.GetResults, this.ResultType, target_shapeids, page);
            this.WriteObject(dt);

            this.WriteVerboseEx("End Query");
        }

        private void addcell(VisioAutomation.ShapeSheet.Query.CellQuery q, bool b, string name)
        {
            var dic = this.GetShapeCellDictionary();
            if (b)
            {
                q.Columns.Add(dic[name], name);
            }
        }

        private static Dictionary<string, VA.ShapeSheet.SRC> dic_cellname_to_src;

        private Dictionary<string, VA.ShapeSheet.SRC> GetShapeCellDictionary()
        {
            if (dic_cellname_to_src == null)
            {
                dic_cellname_to_src = new Dictionary<string, VA.ShapeSheet.SRC>();
                dic_cellname_to_src["Angle"] = VA.ShapeSheet.SRCConstants.Angle;
                dic_cellname_to_src["BeginX"] = VA.ShapeSheet.SRCConstants.BeginX;
                dic_cellname_to_src["BeginY"] = VA.ShapeSheet.SRCConstants.BeginY;
                dic_cellname_to_src["CharCase"] = VA.ShapeSheet.SRCConstants.CharCase;
                dic_cellname_to_src["CharColor"] = VA.ShapeSheet.SRCConstants.CharColor;
                dic_cellname_to_src["CharColorTransparency"] = VA.ShapeSheet.SRCConstants.CharColorTrans;
                dic_cellname_to_src["CharFont"] = VA.ShapeSheet.SRCConstants.CharFont;
                dic_cellname_to_src["CharFontScale"] = VA.ShapeSheet.SRCConstants.CharFontScale;
                dic_cellname_to_src["CharLetterspace"] = VA.ShapeSheet.SRCConstants.CharLetterspace;
                dic_cellname_to_src["CharSize"] = VA.ShapeSheet.SRCConstants.CharSize;
                dic_cellname_to_src["CharStyle"] = VA.ShapeSheet.SRCConstants.CharStyle;
                dic_cellname_to_src["EndX"] = VA.ShapeSheet.SRCConstants.EndX;
                dic_cellname_to_src["EndY"] = VA.ShapeSheet.SRCConstants.EndY;
                dic_cellname_to_src["FillBkgnd"] = VA.ShapeSheet.SRCConstants.FillBkgnd;
                dic_cellname_to_src["FillBkgndTrans"] = VA.ShapeSheet.SRCConstants.FillBkgndTrans;
                dic_cellname_to_src["FillForegnd"] = VA.ShapeSheet.SRCConstants.FillForegnd;
                dic_cellname_to_src["FillForegndTrans"] = VA.ShapeSheet.SRCConstants.FillForegndTrans;
                dic_cellname_to_src["FillPattern"] = VA.ShapeSheet.SRCConstants.FillPattern;
                dic_cellname_to_src["Height"] = VA.ShapeSheet.SRCConstants.Height;
                dic_cellname_to_src["LineCap"] = VA.ShapeSheet.SRCConstants.LineCap;
                dic_cellname_to_src["LineColor"] = VA.ShapeSheet.SRCConstants.LineColor;
                dic_cellname_to_src["LinePattern"] = VA.ShapeSheet.SRCConstants.LinePattern;
                dic_cellname_to_src["LineWeight"] = VA.ShapeSheet.SRCConstants.LineWeight;
                dic_cellname_to_src["LockAspect"] = VA.ShapeSheet.SRCConstants.LockAspect;
                dic_cellname_to_src["LockBegin"] = VA.ShapeSheet.SRCConstants.LockBegin;
                dic_cellname_to_src["LockCalcWH"] = VA.ShapeSheet.SRCConstants.LockCalcWH;
                dic_cellname_to_src["LockCrop"] = VA.ShapeSheet.SRCConstants.LockCrop;
                dic_cellname_to_src["LockCustProp"] = VA.ShapeSheet.SRCConstants.LockCustProp;
                dic_cellname_to_src["LockDelete"] = VA.ShapeSheet.SRCConstants.LockDelete;
                dic_cellname_to_src["LockEnd"] = VA.ShapeSheet.SRCConstants.LockEnd;
                dic_cellname_to_src["LockFormat"] = VA.ShapeSheet.SRCConstants.LockFormat;
                dic_cellname_to_src["LockFromGroupFormat"] = VA.ShapeSheet.SRCConstants.LockFromGroupFormat;
                dic_cellname_to_src["LockGroup"] = VA.ShapeSheet.SRCConstants.LockGroup;
                dic_cellname_to_src["LockHeight"] = VA.ShapeSheet.SRCConstants.LockHeight;
                dic_cellname_to_src["LockMoveX"] = VA.ShapeSheet.SRCConstants.LockMoveX;
                dic_cellname_to_src["LockMoveY"] = VA.ShapeSheet.SRCConstants.LockMoveY;
                dic_cellname_to_src["LockRotate"] = VA.ShapeSheet.SRCConstants.LockRotate;
                dic_cellname_to_src["LockSelect"] = VA.ShapeSheet.SRCConstants.LockSelect;
                dic_cellname_to_src["LockTextEdit"] = VA.ShapeSheet.SRCConstants.LockTextEdit;
                dic_cellname_to_src["LockThemeColors"] = VA.ShapeSheet.SRCConstants.LockThemeColors;
                dic_cellname_to_src["LockThemeEffects"] = VA.ShapeSheet.SRCConstants.LockThemeEffects;
                dic_cellname_to_src["LockVtxEdit"] = VA.ShapeSheet.SRCConstants.LockVtxEdit;
                dic_cellname_to_src["LockWidth"] = VA.ShapeSheet.SRCConstants.LockWidth;
                dic_cellname_to_src["LocPinX"] = VA.ShapeSheet.SRCConstants.LocPinX;
                dic_cellname_to_src["LocPinY"] = VA.ShapeSheet.SRCConstants.LocPinY;
                dic_cellname_to_src["PinX"] = VA.ShapeSheet.SRCConstants.PinX;
                dic_cellname_to_src["PinY"] = VA.ShapeSheet.SRCConstants.PinY;
                dic_cellname_to_src["Rounding"] = VA.ShapeSheet.SRCConstants.Rounding;
                dic_cellname_to_src["SelectMode"] = VA.ShapeSheet.SRCConstants.SelectMode;
                dic_cellname_to_src["ShdwBkgnd"] = VA.ShapeSheet.SRCConstants.ShdwBkgnd;
                dic_cellname_to_src["ShdwBkgndTrans"] = VA.ShapeSheet.SRCConstants.ShdwBkgndTrans;
                dic_cellname_to_src["ShdwForegnd"] = VA.ShapeSheet.SRCConstants.ShdwForegnd;
                dic_cellname_to_src["ShdwForegndTrans"] = VA.ShapeSheet.SRCConstants.ShdwForegndTrans;
                dic_cellname_to_src["ShdwObliqueAngle"] = VA.ShapeSheet.SRCConstants.ShdwObliqueAngle;
                dic_cellname_to_src["ShdwOffsetX"] = VA.ShapeSheet.SRCConstants.ShdwOffsetX;
                dic_cellname_to_src["ShdwOffsetY"] = VA.ShapeSheet.SRCConstants.ShdwOffsetY;
                dic_cellname_to_src["ShdwPattern"] = VA.ShapeSheet.SRCConstants.ShdwPattern;
                dic_cellname_to_src["ShdwScaleFactor"] = VA.ShapeSheet.SRCConstants.ShdwScaleFactor;
                dic_cellname_to_src["ShdwType"] = VA.ShapeSheet.SRCConstants.ShdwType;
                dic_cellname_to_src["TxtAngle"] = VA.ShapeSheet.SRCConstants.TxtAngle;
                dic_cellname_to_src["TxtHeight"] = VA.ShapeSheet.SRCConstants.TxtHeight;
                dic_cellname_to_src["TxtLocPinX"] = VA.ShapeSheet.SRCConstants.TxtLocPinX;
                dic_cellname_to_src["TxtLocPinY"] = VA.ShapeSheet.SRCConstants.TxtLocPinY;
                dic_cellname_to_src["TxtPinX"] = VA.ShapeSheet.SRCConstants.TxtPinX;
                dic_cellname_to_src["TxtPinY"] = VA.ShapeSheet.SRCConstants.TxtPinY;
                dic_cellname_to_src["TxtWidth"] = VA.ShapeSheet.SRCConstants.TxtWidth;
                dic_cellname_to_src["Width"] = VA.ShapeSheet.SRCConstants.Width;

                dic_cellname_to_src["BeginArrow"] = VA.ShapeSheet.SRCConstants.BeginArrow;
                dic_cellname_to_src["BeginArrowSize"] = VA.ShapeSheet.SRCConstants.BeginArrowSize;
                dic_cellname_to_src["EndArrow"] = VA.ShapeSheet.SRCConstants.EndArrow;
                dic_cellname_to_src["EndArrowSize"] = VA.ShapeSheet.SRCConstants.EndArrowSize;

                dic_cellname_to_src["HideText"] = VA.ShapeSheet.SRCConstants.HideText;
            }
            return dic_cellname_to_src;
        }
    }

    /*

Angle  
BeginArrow  
BeginArrowSize  
BeginX  
BeginY  
CharCase  
CharColor  
CharColorTransparency  
CharFont  
CharFontScale  
CharLetterspace  
CharSize  
CharStyle  
EndArrow  
EndArrowSize  
EndX  
EndY  
FillBkgnd  
FillBkgndTrans  
FillForegnd  
FillForegndTrans  
FillPattern  
Height  
HideText  
LineCap  
LineColor  
LinePattern  
LineWeight  
LockAspect  
LockBegin  
LockCalcWH  
LockCrop  
LockCustProp  
LockDelete  
LockEnd  
LockFormat  
LockFromGroupFormat  
LockGroup  
LockHeight  
LockMoveX  
LockMoveY  
LockRotate  
LockSelect  
LockTextEdit  
LockThemeColors  
LockThemeEffects  
LockVtxEdit  
LockWidth  
LocPinX  
LocPinY  
PinX  
PinY  
Rounding  
SelectMode  
ShdwBkgnd  
ShdwBkgndTrans  
ShdwForegnd  
ShdwForegndTrans  
ShdwObliqueAngle  
ShdwOffsetX  
ShdwOffsetY  
ShdwPattern  
ShdwScalefactor  
ShdwType  
TxtAngle  
TxtHeight  
TxtLocPinX  
TxtLocPinY  
TxtPinX  
TxtPinY  
TxtWidth  
Width  

     
     */
}