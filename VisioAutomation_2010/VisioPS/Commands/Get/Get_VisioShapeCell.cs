using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using System.Linq;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioShapeCell")]
    public class Get_VisioShapeCell: VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = true,Position=0)]
        [SMA.ValidateSet( 
            "Angle", "BeginArrow", "BeginArrowSize", "BeginX", "BeginY", "CharCase", "CharColor", "CharColorTransparency", "CharFont",
            "CharFontScale", "CharLetterspace", "CharSize", "CharStyle", "EndArrow", "EndArrowSize", "EndX", "EndY", "FillBkgnd",
            "FillBkgndTrans", "FillForegnd", "FillForegndTrans", "FillPattern", 
            "Height", "LineCap", "LineColor", "LinePattern", "LineWeight", "LockAspect", "LockBegin", 
            "LockCalcWH", "LockCrop", "LockCustProp", "LockDelete", "LockEnd", "LockFormat", "LockFromGroupFormat", 
            "LockGroup", "LockHeight", "LockMoveX", "LockMoveY", "LockRotate", "LockSelect", "LockTextEdit", 
            "LockThemeColors", "LockThemeEffects", "LockVtxEdit", "LockWidth", "LocPinX", "LocPinY", "PinX",
            "PinY", "Rounding", "SelectMode", "ShdwBkgnd", "ShdwBkgndTrans", "ShdwForegnd",
            "ShdwForegndTrans", "ShdwObliqueAngle", "ShdwOffsetX", "ShdwOffsetY", "ShdwPattern",
            "ShdwScalefactor", "ShdwType", "TxtAngle", "TxtHeight", "TxtLocPinX", "TxtLocPinY", "TxtPinX", 
            "TxtPinY", "TxtWidth", "Width", "HideText")]
        public string[] Cells { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter GetResults;

        [SMA.Parameter(Mandatory = false)]
        public ResultType ResultType = ResultType.String;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            var query = new VisioAutomation.ShapeSheet.Query.QueryEx();

            var target_shapes = this.Shapes ?? scriptingsession.Selection.GetShapes();
            var target_shapeids = target_shapes.Select(s => s.ID).ToList();

            var dic = GetShapeCellDictionary();
            foreach (var cell in this.Cells)
            {
                query.AddColumn(dic[cell], cell);   
            }

            var page = scriptingsession.Page.Get();

            this.WriteVerboseEx("Number of Shapes : {0}", target_shapes.Count);
            this.WriteVerboseEx("Number of Cells: {0}", query.Cells.Count);

            this.WriteVerboseEx("Start Query");

            var dt = VisioPSUtil.QueryToDataTable(query, this.GetResults, this.ResultType, target_shapeids, page);
            this.WriteObject(dt);

            this.WriteVerboseEx("End Query");
        }

        private static Dictionary<string, VA.ShapeSheet.SRC> dic_cellname_to_src;
        
        private Dictionary<string, VA.ShapeSheet.SRC> GetShapeCellDictionary()
        {
            if (dic_cellname_to_src == null)
            {
                dic_cellname_to_src = new Dictionary<string, VA.ShapeSheet.SRC>(this.Cells.Count());
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
}