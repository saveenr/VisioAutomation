using System.Collections.Generic;
using VisioAutomation.Scripting;
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
        [SMA.ValidateSet("Angle", "BeginX", "BeginY", "CharCase", "CharColor", "CharColorTransparency", "CharFont", "CharFontScale", "CharLetterspace", "CharSize", "CharStyle", "EndX", "EndY", "FillBackgroundColor", "FillBackgroundtransparency", "FillForegroundColor", "FillForegroundtransparency", "FillPattern", "Height", "LineCap", "LineColor", "LinePattern", "LineWeight", "LockAspect", "LockBegin", "LockCalcWH", "LockCrop", "LockCustProp", "LockDelete", "LockEnd", "LockFormat", "LockFromGroupFormat", "LockGroup", "LockHeight", "LockMoveX", "LockMoveY", "LockRotate", "LockSelect", "LockTextEdit", "LockThemeColors", "LockThemeEffects", "LockVtxEdit", "LockWidth", "LocPinX", "LocPinY", "PinX", "PinY", "Rounding", "SelectMode", "ShadowBackground", "ShadowBackgroundTransparency", "ShadowForeground", "ShadowForegroundTransparency", "ShadowObliqueAngle", "ShadowOffsetX", "ShadowOffsetY", "ShadowPattern", "ShadowScalefactor", "ShadowType", "TxtAngle", "TxtHeight", "TxtLocPinX", "TxtLocPinY", "TxtPinX", "TxtPinY", "TxtWidth", "Width")]
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

            var query = new VisioAutomation.ShapeSheet.Query.CellQuery();

            var target_shapes = this.Shapes ?? scriptingsession.Selection.GetShapes();
            var target_shapeids = target_shapes.Select(s => s.ID).ToList();

            var dic = new Dictionary<string,VA.ShapeSheet.SRC>(this.Cells.Count());
            dic["Angle"] = VA.ShapeSheet.SRCConstants.Angle;
            dic["BeginX"] = VA.ShapeSheet.SRCConstants.BeginX;
            dic["BeginY"] = VA.ShapeSheet.SRCConstants.BeginY;
            dic["CharCase"] = VA.ShapeSheet.SRCConstants.Char_Case;
            dic["CharColor"] = VA.ShapeSheet.SRCConstants.Char_Color;
            dic["CharColorTransparency"] = VA.ShapeSheet.SRCConstants.Char_ColorTrans;
            dic["CharFont"] = VA.ShapeSheet.SRCConstants.Char_Font;
            dic["CharFontScale"] = VA.ShapeSheet.SRCConstants.Char_FontScale;
            dic["CharLetterspace"] = VA.ShapeSheet.SRCConstants.Char_Letterspace;
            dic["CharSize"] = VA.ShapeSheet.SRCConstants.Char_Size;
            dic["CharStyle"] = VA.ShapeSheet.SRCConstants.Char_Style;
            dic["EndX"] = VA.ShapeSheet.SRCConstants.EndX;
            dic["EndY"] = VA.ShapeSheet.SRCConstants.EndY;
            dic["FillBackgroundColor"] = VA.ShapeSheet.SRCConstants.FillBkgnd;
            dic["FillBackgroundtransparency"] = VA.ShapeSheet.SRCConstants.FillBkgndTrans;
            dic["FillForegroundColor"] = VA.ShapeSheet.SRCConstants.FillForegnd;
            dic["FillForegroundtransparency"] = VA.ShapeSheet.SRCConstants.FillForegndTrans;
            dic["FillPattern"] = VA.ShapeSheet.SRCConstants.FillPattern;
            dic["Height"] = VA.ShapeSheet.SRCConstants.Height;
            dic["LineCap"] = VA.ShapeSheet.SRCConstants.LineCap;
            dic["LineColor"] = VA.ShapeSheet.SRCConstants.LineColor;
            dic["LinePattern"] = VA.ShapeSheet.SRCConstants.LinePattern;
            dic["LineWeight"] = VA.ShapeSheet.SRCConstants.LineWeight;
            dic["LockAspect"] = VA.ShapeSheet.SRCConstants.LockAspect;
            dic["LockBegin"] = VA.ShapeSheet.SRCConstants.LockBegin;
            dic["LockCalcWH"] = VA.ShapeSheet.SRCConstants.LockCalcWH;
            dic["LockCrop"] = VA.ShapeSheet.SRCConstants.LockCrop;
            dic["LockCustProp"] = VA.ShapeSheet.SRCConstants.LockCustProp;
            dic["LockDelete"] = VA.ShapeSheet.SRCConstants.LockDelete;
            dic["LockEnd"] = VA.ShapeSheet.SRCConstants.LockEnd;
            dic["LockFormat"] = VA.ShapeSheet.SRCConstants.LockFormat;
            dic["LockFromGroupFormat"] = VA.ShapeSheet.SRCConstants.LockFromGroupFormat;
            dic["LockGroup"] = VA.ShapeSheet.SRCConstants.LockGroup;
            dic["LockHeight"] = VA.ShapeSheet.SRCConstants.LockHeight;
            dic["LockMoveX"] = VA.ShapeSheet.SRCConstants.LockMoveX;
            dic["LockMoveY"] = VA.ShapeSheet.SRCConstants.LockMoveY;
            dic["LockRotate"] = VA.ShapeSheet.SRCConstants.LockRotate;
            dic["LockSelect"] = VA.ShapeSheet.SRCConstants.LockSelect;
            dic["LockTextEdit"] = VA.ShapeSheet.SRCConstants.LockTextEdit;
            dic["LockThemeColors"] = VA.ShapeSheet.SRCConstants.LockThemeColors;
            dic["LockThemeEffects"] = VA.ShapeSheet.SRCConstants.LockThemeEffects;
            dic["LockVtxEdit"] = VA.ShapeSheet.SRCConstants.LockVtxEdit;
            dic["LockWidth"] = VA.ShapeSheet.SRCConstants.LockWidth;
            dic["LocPinX"] = VA.ShapeSheet.SRCConstants.LocPinX;
            dic["LocPinY"] = VA.ShapeSheet.SRCConstants.LocPinY;
            dic["PinX"] = VA.ShapeSheet.SRCConstants.PinX;
            dic["PinY"] = VA.ShapeSheet.SRCConstants.PinY;
            dic["Rounding"] = VA.ShapeSheet.SRCConstants.Rounding;
            dic["SelectMode"] = VA.ShapeSheet.SRCConstants.SelectMode;
            dic["ShadowBackground"] = VA.ShapeSheet.SRCConstants.ShdwBkgnd;
            dic["ShadowBackgroundTransparency"] = VA.ShapeSheet.SRCConstants.ShdwBkgndTrans;
            dic["ShadowForeground"] = VA.ShapeSheet.SRCConstants.ShdwForegnd;
            dic["ShadowForegroundTransparency"] = VA.ShapeSheet.SRCConstants.ShdwForegndTrans;
            dic["ShadowObliqueAngle"] = VA.ShapeSheet.SRCConstants.ShdwObliqueAngle;
            dic["ShadowOffsetX"] = VA.ShapeSheet.SRCConstants.ShdwOffsetX;
            dic["ShadowOffsetY"] = VA.ShapeSheet.SRCConstants.ShdwOffsetY;
            dic["ShadowPattern"] = VA.ShapeSheet.SRCConstants.ShdwPattern;
            dic["ShadowScalefactor"] = VA.ShapeSheet.SRCConstants.ShdwScaleFactor;
            dic["ShadowType"] = VA.ShapeSheet.SRCConstants.ShdwType;
            dic["TxtAngle"] = VA.ShapeSheet.SRCConstants.TxtAngle;
            dic["TxtHeight"] = VA.ShapeSheet.SRCConstants.TxtHeight;
            dic["TxtLocPinX"] = VA.ShapeSheet.SRCConstants.TxtLocPinX;
            dic["TxtLocPinY"] = VA.ShapeSheet.SRCConstants.TxtLocPinY;
            dic["TxtPinX"] = VA.ShapeSheet.SRCConstants.TxtPinX;
            dic["TxtPinY"] = VA.ShapeSheet.SRCConstants.TxtPinY;
            dic["TxtWidth"] = VA.ShapeSheet.SRCConstants.TxtWidth;
            dic["Width"] = VA.ShapeSheet.SRCConstants.Width;


            foreach (var cell in this.Cells)
            {
                query.AddColumn(dic[cell], cell);   
            }

            var page = scriptingsession.Page.Get();

            this.WriteVerboseEx("Number of Shapes : {0}", target_shapes.Count);
            this.WriteVerboseEx("Number of Cells: {0}", query.Columns.Count);

            this.WriteVerboseEx("Start Query");

            var names = query.Columns.Select(c => c.Name).ToList();
            if (this.GetResults)
            {
                if (this.ResultType == ResultType.String)
                {
                    var output = query.GetResults<string>(page, target_shapeids);
                    this.WriteObject(todatatable(output,names));
                }
                else if (this.ResultType == ResultType.Boolean)
                {
                    var output = query.GetResults<bool>(page, target_shapeids);
                    this.WriteObject(todatatable(output, names));
                }
                else if (this.ResultType == ResultType.Double)
                {
                    var output = query.GetResults<double>(page, target_shapeids);
                    this.WriteObject(todatatable(output, names));
                }
                else if (this.ResultType == ResultType.Integer)
                {
                    var output = query.GetResults<int>(page, target_shapeids);
                    this.WriteObject(todatatable(output, names));
                }
                else
                {
                    throw new VisioApplicationException("Unsupported Result type");
                }

            }
            else
            {
                var output = query.GetFormulas(page, target_shapeids);
                this.WriteObject(todatatable(output, names));
            }

            this.WriteVerboseEx("End Query");
        }

        private System.Data.DataTable todatatable<T>(VA.ShapeSheet.Data.Table<T> output, IList<string> names )
        {
            var dt = new System.Data.DataTable();
            foreach (string name in names)
            {
                dt.Columns.Add(name, typeof (T));
            }
            int colcount = names.Count;
            var arr = new object[colcount];
            for (int r = 0; r < output.RowCount; r++)
            {
                for (int i = 0; i < colcount; i++)
                {
                    arr[i] = output[r, i];
                }
                dt.Rows.Add(arr);
            }
            return dt;
        }
    }
}