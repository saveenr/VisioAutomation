using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using System.Linq;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioShapeCell")]
    public class Get_VisioShapeCell : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false, Position = 0)]
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
            var query = new VisioAutomation.ShapeSheet.Query.CellQuery();

            var target_shapes = this.Shapes ?? this.client.Selection.GetShapes();
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

            var dic = GetShapeCellDictionary();
            Get_VisioPageCell.SetFromCellNames(query, this.Cells, dic);

            var surface = this.client.Draw.GetDrawingSurfaceSafe();

            this.WriteVerbose("Number of Shapes : {0}", target_shapes.Count);
            this.WriteVerbose("Number of Cells: {0}", query.Columns.Count);

            this.WriteVerbose("Start Query");

            var dt = Helpers.QueryToDataTable(query, this.GetResults, this.ResultType, target_shapeids, surface);
            this.WriteObject(dt);

            this.WriteVerbose("End Query");
        }

        private void addcell(VisioAutomation.ShapeSheet.Query.CellQuery q, bool b, string name)
        {
            var dic = GetShapeCellDictionary();
            if (b)
            {
                q.Columns.Add(dic[name], name);
            }
        }

        private static CellMap callmap;

        public static CellMap GetShapeCellDictionary()
        {
            if (callmap == null)
            {
                callmap = new CellMap();
                callmap["Angle"] = VA.ShapeSheet.SRCConstants.Angle;
                callmap["BeginX"] = VA.ShapeSheet.SRCConstants.BeginX;
                callmap["BeginY"] = VA.ShapeSheet.SRCConstants.BeginY;
                callmap["CharCase"] = VA.ShapeSheet.SRCConstants.CharCase;
                callmap["CharColor"] = VA.ShapeSheet.SRCConstants.CharColor;
                callmap["CharColorTransparency"] = VA.ShapeSheet.SRCConstants.CharColorTrans;
                callmap["CharFont"] = VA.ShapeSheet.SRCConstants.CharFont;
                callmap["CharFontScale"] = VA.ShapeSheet.SRCConstants.CharFontScale;
                callmap["CharLetterspace"] = VA.ShapeSheet.SRCConstants.CharLetterspace;
                callmap["CharSize"] = VA.ShapeSheet.SRCConstants.CharSize;
                callmap["CharStyle"] = VA.ShapeSheet.SRCConstants.CharStyle;
                callmap["EndX"] = VA.ShapeSheet.SRCConstants.EndX;
                callmap["EndY"] = VA.ShapeSheet.SRCConstants.EndY;
                callmap["FillBkgnd"] = VA.ShapeSheet.SRCConstants.FillBkgnd;
                callmap["FillBkgndTrans"] = VA.ShapeSheet.SRCConstants.FillBkgndTrans;
                callmap["FillForegnd"] = VA.ShapeSheet.SRCConstants.FillForegnd;
                callmap["FillForegndTrans"] = VA.ShapeSheet.SRCConstants.FillForegndTrans;
                callmap["FillPattern"] = VA.ShapeSheet.SRCConstants.FillPattern;
                callmap["Height"] = VA.ShapeSheet.SRCConstants.Height;
                callmap["LineCap"] = VA.ShapeSheet.SRCConstants.LineCap;
                callmap["LineColor"] = VA.ShapeSheet.SRCConstants.LineColor;
                callmap["LinePattern"] = VA.ShapeSheet.SRCConstants.LinePattern;
                callmap["LineWeight"] = VA.ShapeSheet.SRCConstants.LineWeight;
                callmap["LockAspect"] = VA.ShapeSheet.SRCConstants.LockAspect;
                callmap["LockBegin"] = VA.ShapeSheet.SRCConstants.LockBegin;
                callmap["LockCalcWH"] = VA.ShapeSheet.SRCConstants.LockCalcWH;
                callmap["LockCrop"] = VA.ShapeSheet.SRCConstants.LockCrop;
                callmap["LockCustProp"] = VA.ShapeSheet.SRCConstants.LockCustProp;
                callmap["LockDelete"] = VA.ShapeSheet.SRCConstants.LockDelete;
                callmap["LockEnd"] = VA.ShapeSheet.SRCConstants.LockEnd;
                callmap["LockFormat"] = VA.ShapeSheet.SRCConstants.LockFormat;
                callmap["LockFromGroupFormat"] = VA.ShapeSheet.SRCConstants.LockFromGroupFormat;
                callmap["LockGroup"] = VA.ShapeSheet.SRCConstants.LockGroup;
                callmap["LockHeight"] = VA.ShapeSheet.SRCConstants.LockHeight;
                callmap["LockMoveX"] = VA.ShapeSheet.SRCConstants.LockMoveX;
                callmap["LockMoveY"] = VA.ShapeSheet.SRCConstants.LockMoveY;
                callmap["LockRotate"] = VA.ShapeSheet.SRCConstants.LockRotate;
                callmap["LockSelect"] = VA.ShapeSheet.SRCConstants.LockSelect;
                callmap["LockTextEdit"] = VA.ShapeSheet.SRCConstants.LockTextEdit;
                callmap["LockThemeColors"] = VA.ShapeSheet.SRCConstants.LockThemeColors;
                callmap["LockThemeEffects"] = VA.ShapeSheet.SRCConstants.LockThemeEffects;
                callmap["LockVtxEdit"] = VA.ShapeSheet.SRCConstants.LockVtxEdit;
                callmap["LockWidth"] = VA.ShapeSheet.SRCConstants.LockWidth;
                callmap["LocPinX"] = VA.ShapeSheet.SRCConstants.LocPinX;
                callmap["LocPinY"] = VA.ShapeSheet.SRCConstants.LocPinY;
                callmap["PinX"] = VA.ShapeSheet.SRCConstants.PinX;
                callmap["PinY"] = VA.ShapeSheet.SRCConstants.PinY;
                callmap["Rounding"] = VA.ShapeSheet.SRCConstants.Rounding;
                callmap["SelectMode"] = VA.ShapeSheet.SRCConstants.SelectMode;
                callmap["ShdwBkgnd"] = VA.ShapeSheet.SRCConstants.ShdwBkgnd;
                callmap["ShdwBkgndTrans"] = VA.ShapeSheet.SRCConstants.ShdwBkgndTrans;
                callmap["ShdwForegnd"] = VA.ShapeSheet.SRCConstants.ShdwForegnd;
                callmap["ShdwForegndTrans"] = VA.ShapeSheet.SRCConstants.ShdwForegndTrans;
                callmap["ShdwObliqueAngle"] = VA.ShapeSheet.SRCConstants.ShdwObliqueAngle;
                callmap["ShdwOffsetX"] = VA.ShapeSheet.SRCConstants.ShdwOffsetX;
                callmap["ShdwOffsetY"] = VA.ShapeSheet.SRCConstants.ShdwOffsetY;
                callmap["ShdwPattern"] = VA.ShapeSheet.SRCConstants.ShdwPattern;
                callmap["ShdwScaleFactor"] = VA.ShapeSheet.SRCConstants.ShdwScaleFactor;
                callmap["ShdwType"] = VA.ShapeSheet.SRCConstants.ShdwType;
                callmap["TxtAngle"] = VA.ShapeSheet.SRCConstants.TxtAngle;
                callmap["TxtHeight"] = VA.ShapeSheet.SRCConstants.TxtHeight;
                callmap["TxtLocPinX"] = VA.ShapeSheet.SRCConstants.TxtLocPinX;
                callmap["TxtLocPinY"] = VA.ShapeSheet.SRCConstants.TxtLocPinY;
                callmap["TxtPinX"] = VA.ShapeSheet.SRCConstants.TxtPinX;
                callmap["TxtPinY"] = VA.ShapeSheet.SRCConstants.TxtPinY;
                callmap["TxtWidth"] = VA.ShapeSheet.SRCConstants.TxtWidth;
                callmap["Width"] = VA.ShapeSheet.SRCConstants.Width;

                callmap["BeginArrow"] = VA.ShapeSheet.SRCConstants.BeginArrow;
                callmap["BeginArrowSize"] = VA.ShapeSheet.SRCConstants.BeginArrowSize;
                callmap["EndArrow"] = VA.ShapeSheet.SRCConstants.EndArrow;
                callmap["EndArrowSize"] = VA.ShapeSheet.SRCConstants.EndArrowSize;

                callmap["HideText"] = VA.ShapeSheet.SRCConstants.HideText;
            }
            return callmap;
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