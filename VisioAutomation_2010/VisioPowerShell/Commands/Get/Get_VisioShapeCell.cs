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

            var surface = this.client.ShapeSheet.GetShapeSheetSurface();

            this.WriteVerbose("Number of Shapes : {0}", target_shapes.Count);
            this.WriteVerbose("Number of Cells: {0}", query.CellColumns.Count);

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
                q.AddCell(dic[name], name);
            }
        }

        private static CellMap map_name_to_cell;

        public static CellMap GetShapeCellDictionary()
        {
            if (map_name_to_cell == null)
            {
                map_name_to_cell = new CellMap();
                map_name_to_cell["Angle"] = VA.ShapeSheet.SRCConstants.Angle;
                map_name_to_cell["BeginX"] = VA.ShapeSheet.SRCConstants.BeginX;
                map_name_to_cell["BeginY"] = VA.ShapeSheet.SRCConstants.BeginY;
                map_name_to_cell["CharCase"] = VA.ShapeSheet.SRCConstants.CharCase;
                map_name_to_cell["CharColor"] = VA.ShapeSheet.SRCConstants.CharColor;
                map_name_to_cell["CharColorTransparency"] = VA.ShapeSheet.SRCConstants.CharColorTrans;
                map_name_to_cell["CharFont"] = VA.ShapeSheet.SRCConstants.CharFont;
                map_name_to_cell["CharFontScale"] = VA.ShapeSheet.SRCConstants.CharFontScale;
                map_name_to_cell["CharLetterspace"] = VA.ShapeSheet.SRCConstants.CharLetterspace;
                map_name_to_cell["CharSize"] = VA.ShapeSheet.SRCConstants.CharSize;
                map_name_to_cell["CharStyle"] = VA.ShapeSheet.SRCConstants.CharStyle;
                map_name_to_cell["EndX"] = VA.ShapeSheet.SRCConstants.EndX;
                map_name_to_cell["EndY"] = VA.ShapeSheet.SRCConstants.EndY;
                map_name_to_cell["FillBkgnd"] = VA.ShapeSheet.SRCConstants.FillBkgnd;
                map_name_to_cell["FillBkgndTrans"] = VA.ShapeSheet.SRCConstants.FillBkgndTrans;
                map_name_to_cell["FillForegnd"] = VA.ShapeSheet.SRCConstants.FillForegnd;
                map_name_to_cell["FillForegndTrans"] = VA.ShapeSheet.SRCConstants.FillForegndTrans;
                map_name_to_cell["FillPattern"] = VA.ShapeSheet.SRCConstants.FillPattern;
                map_name_to_cell["Height"] = VA.ShapeSheet.SRCConstants.Height;
                map_name_to_cell["LineCap"] = VA.ShapeSheet.SRCConstants.LineCap;
                map_name_to_cell["LineColor"] = VA.ShapeSheet.SRCConstants.LineColor;
                map_name_to_cell["LinePattern"] = VA.ShapeSheet.SRCConstants.LinePattern;
                map_name_to_cell["LineWeight"] = VA.ShapeSheet.SRCConstants.LineWeight;
                map_name_to_cell["LockAspect"] = VA.ShapeSheet.SRCConstants.LockAspect;
                map_name_to_cell["LockBegin"] = VA.ShapeSheet.SRCConstants.LockBegin;
                map_name_to_cell["LockCalcWH"] = VA.ShapeSheet.SRCConstants.LockCalcWH;
                map_name_to_cell["LockCrop"] = VA.ShapeSheet.SRCConstants.LockCrop;
                map_name_to_cell["LockCustProp"] = VA.ShapeSheet.SRCConstants.LockCustProp;
                map_name_to_cell["LockDelete"] = VA.ShapeSheet.SRCConstants.LockDelete;
                map_name_to_cell["LockEnd"] = VA.ShapeSheet.SRCConstants.LockEnd;
                map_name_to_cell["LockFormat"] = VA.ShapeSheet.SRCConstants.LockFormat;
                map_name_to_cell["LockFromGroupFormat"] = VA.ShapeSheet.SRCConstants.LockFromGroupFormat;
                map_name_to_cell["LockGroup"] = VA.ShapeSheet.SRCConstants.LockGroup;
                map_name_to_cell["LockHeight"] = VA.ShapeSheet.SRCConstants.LockHeight;
                map_name_to_cell["LockMoveX"] = VA.ShapeSheet.SRCConstants.LockMoveX;
                map_name_to_cell["LockMoveY"] = VA.ShapeSheet.SRCConstants.LockMoveY;
                map_name_to_cell["LockRotate"] = VA.ShapeSheet.SRCConstants.LockRotate;
                map_name_to_cell["LockSelect"] = VA.ShapeSheet.SRCConstants.LockSelect;
                map_name_to_cell["LockTextEdit"] = VA.ShapeSheet.SRCConstants.LockTextEdit;
                map_name_to_cell["LockThemeColors"] = VA.ShapeSheet.SRCConstants.LockThemeColors;
                map_name_to_cell["LockThemeEffects"] = VA.ShapeSheet.SRCConstants.LockThemeEffects;
                map_name_to_cell["LockVtxEdit"] = VA.ShapeSheet.SRCConstants.LockVtxEdit;
                map_name_to_cell["LockWidth"] = VA.ShapeSheet.SRCConstants.LockWidth;
                map_name_to_cell["LocPinX"] = VA.ShapeSheet.SRCConstants.LocPinX;
                map_name_to_cell["LocPinY"] = VA.ShapeSheet.SRCConstants.LocPinY;
                map_name_to_cell["PinX"] = VA.ShapeSheet.SRCConstants.PinX;
                map_name_to_cell["PinY"] = VA.ShapeSheet.SRCConstants.PinY;
                map_name_to_cell["Rounding"] = VA.ShapeSheet.SRCConstants.Rounding;
                map_name_to_cell["SelectMode"] = VA.ShapeSheet.SRCConstants.SelectMode;
                map_name_to_cell["ShdwBkgnd"] = VA.ShapeSheet.SRCConstants.ShdwBkgnd;
                map_name_to_cell["ShdwBkgndTrans"] = VA.ShapeSheet.SRCConstants.ShdwBkgndTrans;
                map_name_to_cell["ShdwForegnd"] = VA.ShapeSheet.SRCConstants.ShdwForegnd;
                map_name_to_cell["ShdwForegndTrans"] = VA.ShapeSheet.SRCConstants.ShdwForegndTrans;
                map_name_to_cell["ShdwObliqueAngle"] = VA.ShapeSheet.SRCConstants.ShdwObliqueAngle;
                map_name_to_cell["ShdwOffsetX"] = VA.ShapeSheet.SRCConstants.ShdwOffsetX;
                map_name_to_cell["ShdwOffsetY"] = VA.ShapeSheet.SRCConstants.ShdwOffsetY;
                map_name_to_cell["ShdwPattern"] = VA.ShapeSheet.SRCConstants.ShdwPattern;
                map_name_to_cell["ShdwScaleFactor"] = VA.ShapeSheet.SRCConstants.ShdwScaleFactor;
                map_name_to_cell["ShdwType"] = VA.ShapeSheet.SRCConstants.ShdwType;
                map_name_to_cell["TxtAngle"] = VA.ShapeSheet.SRCConstants.TxtAngle;
                map_name_to_cell["TxtHeight"] = VA.ShapeSheet.SRCConstants.TxtHeight;
                map_name_to_cell["TxtLocPinX"] = VA.ShapeSheet.SRCConstants.TxtLocPinX;
                map_name_to_cell["TxtLocPinY"] = VA.ShapeSheet.SRCConstants.TxtLocPinY;
                map_name_to_cell["TxtPinX"] = VA.ShapeSheet.SRCConstants.TxtPinX;
                map_name_to_cell["TxtPinY"] = VA.ShapeSheet.SRCConstants.TxtPinY;
                map_name_to_cell["TxtWidth"] = VA.ShapeSheet.SRCConstants.TxtWidth;
                map_name_to_cell["Width"] = VA.ShapeSheet.SRCConstants.Width;

                map_name_to_cell["BeginArrow"] = VA.ShapeSheet.SRCConstants.BeginArrow;
                map_name_to_cell["BeginArrowSize"] = VA.ShapeSheet.SRCConstants.BeginArrowSize;
                map_name_to_cell["EndArrow"] = VA.ShapeSheet.SRCConstants.EndArrow;
                map_name_to_cell["EndArrowSize"] = VA.ShapeSheet.SRCConstants.EndArrowSize;

                map_name_to_cell["HideText"] = VA.ShapeSheet.SRCConstants.HideText;
            }
            return map_name_to_cell;
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