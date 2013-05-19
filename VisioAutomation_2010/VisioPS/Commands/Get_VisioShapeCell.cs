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

            var hs = new HashSet<string>( this.Cells );

            if (hs.Contains("Angle")) { query.AddColumn(VA.ShapeSheet.SRCConstants.Angle, "Angle"); }
            if (hs.Contains("BeginX")) { query.AddColumn(VA.ShapeSheet.SRCConstants.BeginX, "BeginX"); }
            if (hs.Contains("BeginY")) { query.AddColumn(VA.ShapeSheet.SRCConstants.BeginY, "BeginY"); }
            if (hs.Contains("CharCase")) { query.AddColumn(VA.ShapeSheet.SRCConstants.Char_Case, "CharCase"); }
            if (hs.Contains("CharColor")) { query.AddColumn(VA.ShapeSheet.SRCConstants.Char_Color, "CharColor"); }
            if (hs.Contains("CharColorTransparency")) { query.AddColumn(VA.ShapeSheet.SRCConstants.Char_ColorTrans, "CharColorTransparency"); }
            if (hs.Contains("CharFont")) { query.AddColumn(VA.ShapeSheet.SRCConstants.Char_Font, "CharFont"); }
            if (hs.Contains("CharFontScale")) { query.AddColumn(VA.ShapeSheet.SRCConstants.Char_FontScale, "CharFontScale"); }
            if (hs.Contains("CharLetterspace")) { query.AddColumn(VA.ShapeSheet.SRCConstants.Char_Letterspace, "CharLetterspace"); }
            if (hs.Contains("CharSize")) { query.AddColumn(VA.ShapeSheet.SRCConstants.Char_Size, "CharSize"); }
            if (hs.Contains("CharStyle")) { query.AddColumn(VA.ShapeSheet.SRCConstants.Char_Style, "CharStyle"); }
            if (hs.Contains("EndX")) { query.AddColumn(VA.ShapeSheet.SRCConstants.EndX, "EndX"); }
            if (hs.Contains("EndY")) { query.AddColumn(VA.ShapeSheet.SRCConstants.EndY, "EndY"); }
            if (hs.Contains("FillBackgroundColor")) { query.AddColumn(VA.ShapeSheet.SRCConstants.FillBkgnd, "FillBackgroundColor"); }
            if (hs.Contains("FillBackgroundtransparency")) { query.AddColumn(VA.ShapeSheet.SRCConstants.FillBkgndTrans, "FillBackgroundtransparency"); }
            if (hs.Contains("FillForegroundColor")) { query.AddColumn(VA.ShapeSheet.SRCConstants.FillForegnd, "FillForegroundColor"); }
            if (hs.Contains("FillForegroundtransparency")) { query.AddColumn(VA.ShapeSheet.SRCConstants.FillForegndTrans, "FillForegroundtransparency"); }
            if (hs.Contains("FillPattern")) { query.AddColumn(VA.ShapeSheet.SRCConstants.FillPattern, "FillPattern"); }
            if (hs.Contains("Height")) { query.AddColumn(VA.ShapeSheet.SRCConstants.Height, "Height"); }
            if (hs.Contains("LineCap")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LineCap, "LineCap"); }
            if (hs.Contains("LineColor")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LineColor, "LineColor"); }
            if (hs.Contains("LinePattern")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LinePattern, "LinePattern"); }
            if (hs.Contains("LineWeight")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LineWeight, "LineWeight"); }
            if (hs.Contains("LockAspect")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockAspect, "LockAspect"); }
            if (hs.Contains("LockBegin")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockBegin, "LockBegin"); }
            if (hs.Contains("LockCalcWH")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockCalcWH, "LockCalcWH"); }
            if (hs.Contains("LockCrop")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockCrop, "LockCrop"); }
            if (hs.Contains("LockCustProp")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockCustProp, "LockCustProp"); }
            if (hs.Contains("LockDelete")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockDelete, "LockDelete"); }
            if (hs.Contains("LockEnd")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockEnd, "LockEnd"); }
            if (hs.Contains("LockFormat")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockFormat, "LockFormat"); }
            if (hs.Contains("LockFromGroupFormat")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockFromGroupFormat, "LockFromGroupFormat"); }
            if (hs.Contains("LockGroup")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockGroup, "LockGroup"); }
            if (hs.Contains("LockHeight")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockHeight, "LockHeight"); }
            if (hs.Contains("LockMoveX")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockMoveX, "LockMoveX"); }
            if (hs.Contains("LockMoveY")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockMoveY, "LockMoveY"); }
            if (hs.Contains("LockRotate")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockRotate, "LockRotate"); }
            if (hs.Contains("LockSelect")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockSelect, "LockSelect"); }
            if (hs.Contains("LockTextEdit")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockTextEdit, "LockTextEdit"); }
            if (hs.Contains("LockThemeColors")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockThemeColors, "LockThemeColors"); }
            if (hs.Contains("LockThemeEffects")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockThemeEffects, "LockThemeEffects"); }
            if (hs.Contains("LockVtxEdit")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockVtxEdit, "LockVtxEdit"); }
            if (hs.Contains("LockWidth")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LockWidth, "LockWidth"); }
            if (hs.Contains("LocPinX")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LocPinX, "LocPinX"); }
            if (hs.Contains("LocPinY")) { query.AddColumn(VA.ShapeSheet.SRCConstants.LocPinY, "LocPinY"); }
            if (hs.Contains("PinX")) { query.AddColumn(VA.ShapeSheet.SRCConstants.PinX, "PinX"); }
            if (hs.Contains("PinY")) { query.AddColumn(VA.ShapeSheet.SRCConstants.PinY, "PinY"); }
            if (hs.Contains("Rounding")) { query.AddColumn(VA.ShapeSheet.SRCConstants.Rounding, "Rounding"); }
            if (hs.Contains("SelectMode")) { query.AddColumn(VA.ShapeSheet.SRCConstants.SelectMode, "SelectMode"); }
            if (hs.Contains("ShadowBackground")) { query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwBkgnd, "ShadowBackground"); }
            if (hs.Contains("ShadowBackgroundTransparency")) { query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwBkgndTrans, "ShadowBackgroundTransparency"); }
            if (hs.Contains("ShadowForeground")) { query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwForegnd, "ShadowForeground"); }
            if (hs.Contains("ShadowForegroundTransparency")) { query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwForegndTrans, "ShadowForegroundTransparency"); }
            if (hs.Contains("ShadowObliqueAngle")) { query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwObliqueAngle, "ShadowObliqueAngle"); }
            if (hs.Contains("ShadowOffsetX")) { query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwOffsetX, "ShadowOffsetX"); }
            if (hs.Contains("ShadowOffsetY")) { query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwOffsetY, "ShadowOffsetY"); }
            if (hs.Contains("ShadowPattern")) { query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwPattern, "ShadowPattern"); }
            if (hs.Contains("ShadowScalefactor")) { query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwScaleFactor, "ShadowScalefactor"); }
            if (hs.Contains("ShadowType")) { query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwType, "ShadowType"); }
            if (hs.Contains("TxtAngle")) { query.AddColumn(VA.ShapeSheet.SRCConstants.TxtAngle, "TxtAngle"); }
            if (hs.Contains("TxtHeight")) { query.AddColumn(VA.ShapeSheet.SRCConstants.TxtHeight, "TxtHeight"); }
            if (hs.Contains("TxtLocPinX")) { query.AddColumn(VA.ShapeSheet.SRCConstants.TxtLocPinX, "TxtLocPinX"); }
            if (hs.Contains("TxtLocPinY")) { query.AddColumn(VA.ShapeSheet.SRCConstants.TxtLocPinY, "TxtLocPinY"); }
            if (hs.Contains("TxtPinX")) { query.AddColumn(VA.ShapeSheet.SRCConstants.TxtPinX, "TxtPinX"); }
            if (hs.Contains("TxtPinY")) { query.AddColumn(VA.ShapeSheet.SRCConstants.TxtPinY, "TxtPinY"); }
            if (hs.Contains("TxtWidth")) { query.AddColumn(VA.ShapeSheet.SRCConstants.TxtWidth, "TxtWidth"); }
            if (hs.Contains("Width")) { query.AddColumn(VA.ShapeSheet.SRCConstants.Width, "Width"); }

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