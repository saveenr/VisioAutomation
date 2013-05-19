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
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter Width { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter Height { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter PinX { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter PinY { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LocPinX { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LocPinY { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter Angle { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter FillPattern { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter FillForegroundColor { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter FillForegroundtransparency { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter FillBackgroundColor { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter FillBackgroundtransparency { get; set; }       
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LinePattern { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LineWeight{ get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LineColor { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LineCap { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter Rounding { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter CharCase { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter CharColor { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter CharFont { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter CharFontScale { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter CharLetterspace { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter CharSize { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter CharStyle { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter CharColorTransparency { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter BeginX { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter BeginY{ get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter EndX{ get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter EndY{ get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter ShadowBackground { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter ShadowBackgroundTransparency { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter ShadowForeground { get; set; }        
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter ShadowForegroundTransparency { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter ShadowObliqueAngle { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter ShadowOffsetX { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter ShadowOffsetY { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter ShadowPattern { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter ShadowScalefactor { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter ShadowType { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter SelectMode { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockAspect { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockBegin { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockCalcWH { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockCrop { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockCustProp { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockDelete { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockEnd { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockFormat { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockFromGroupFormat { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockGroup { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockHeight { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockMoveX { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockMoveY { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockRotate { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockSelect { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockTextEdit { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockThemeColors { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockThemeEffects { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockVtxEdit { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter LockWidth { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter TxtAngle { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter TxtHeight { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter TxtLocPinX  { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter TxtLocPinY { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter TxtPinX { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter TxtPinY { get; set; }
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter TxtWidth { get; set; }
        
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

            if (this.Width)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.Width, "Width ");
            }
            if (this.Height)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.Height, "Height ");
            }
            if (this.PinX)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.PinX, "PinX ");
            }
            if (this.PinY)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.PinY, "PinY ");
            }
            if (this.LocPinX)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LocPinX, "LocPinX ");
            }
            if (this.LocPinY)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LocPinY, "LocPinY ");
            }
            if (this.Angle)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.Angle, "Angle ");
            }
            if (this.FillPattern)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.FillPattern, "FillPattern ");
            }
            if (this.FillForegroundColor)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.FillForegnd, "FillForegroundColor ");
            }
            if (this.FillForegroundtransparency)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.FillForegndTrans, "FillForegroundtransparency ");
            }
            if (this.FillBackgroundColor)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.FillBkgnd, "FillBackgroundColor ");
            }
            if (this.FillBackgroundtransparency)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.FillBkgndTrans, "FillBackgroundtransparency ");
            }
            if (this.LinePattern)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LinePattern, "LinePattern ");
            }
            if (this.LineWeight)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LineWeight, "LineWeight");
            }
            if (this.LineColor)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LineColor, "LineColor ");
            }
            if (this.LineCap)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LineCap, "LineCap ");
            }
            if (this.Rounding)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.Rounding, "Rounding ");
            }
            if (this.CharCase)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.Char_Case, "CharCase ");
            }
            if (this.CharColor)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.Char_Color, "CharColor ");
            }
            if (this.CharFont)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.Char_Font, "CharFont ");
            }
            if (this.CharFontScale)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.Char_FontScale, "CharFontScale ");
            }
            if (this.CharLetterspace)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.Char_Letterspace, "CharLetterspace ");
            }
            if (this.CharSize)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.Char_Size, "CharSize ");
            }
            if (this.CharStyle)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.Char_Style, "CharStyle ");
            }
            if (this.CharColorTransparency)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.Char_ColorTrans, "CharColorTransparency ");
            }
            if (this.BeginX)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.BeginX, "BeginX ");
            }
            if (this.BeginY)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.BeginY, "BeginY");
            }
            if (this.EndX)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.EndX, "EndX");
            }
            if (this.EndY)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.EndY, "EndY");
            }
            if (this.ShadowBackground)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwBkgnd, "ShadowBackground ");
            }
            if (this.ShadowBackgroundTransparency)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwBkgndTrans, "ShadowBackgroundTransparency ");
            }
            if (this.ShadowForeground)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwForegnd, "ShadowForeground ");
            }
            if (this.ShadowForegroundTransparency)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwForegndTrans, "ShadowForegroundTransparency ");
            }
            if (this.ShadowObliqueAngle)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwObliqueAngle, "ShadowObliqueAngle ");
            }
            if (this.ShadowOffsetX)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwOffsetX, "ShadowOffsetX ");
            }
            if (this.ShadowOffsetY)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwOffsetY, "ShadowOffsetY ");
            }
            if (this.ShadowPattern)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwPattern, "ShadowPattern ");
            }
            if (this.ShadowScalefactor)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwScaleFactor, "ShadowScalefactor ");
            }
            if (this.ShadowType)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.ShdwType, "ShadowType ");
            }
            if (this.SelectMode)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.SelectMode, "SelectMode ");
            }
            if (this.LockAspect)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockAspect, "LockAspect ");
            }
            if (this.LockBegin)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockBegin, "LockBegin ");
            }
            if (this.LockCalcWH)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockCalcWH, "LockCalcWH ");
            }
            if (this.LockCrop)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockCrop, "LockCrop ");
            }
            if (this.LockCustProp)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockCustProp, "LockCustProp ");
            }
            if (this.LockDelete)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockDelete, "LockDelete ");
            }
            if (this.LockEnd)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockEnd, "LockEnd ");
            }
            if (this.LockFormat)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockFormat, "LockFormat ");
            }
            if (this.LockFromGroupFormat)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockFromGroupFormat, "LockFromGroupFormat ");
            }
            if (this.LockGroup)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockGroup, "LockGroup ");
            }
            if (this.LockHeight)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockHeight, "LockHeight ");
            }
            if (this.LockMoveX)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockMoveX, "LockMoveX ");
            }
            if (this.LockMoveY)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockMoveY, "LockMoveY ");
            }
            if (this.LockRotate)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockRotate, "LockRotate ");
            }
            if (this.LockSelect)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockSelect, "LockSelect ");
            }
            if (this.LockTextEdit)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockTextEdit, "LockTextEdit ");
            }
            if (this.LockThemeColors)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockThemeColors, "LockThemeColors ");
            }
            if (this.LockThemeEffects)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockThemeEffects, "LockThemeEffects ");
            }
            if (this.LockVtxEdit)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockVtxEdit, "LockVtxEdit ");
            }
            if (this.LockWidth)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.LockWidth, "LockWidth ");
            }
            if (this.TxtAngle)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.TxtAngle, "TxtAngle ");
            }
            if (this.TxtHeight)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.TxtHeight, "TxtHeight ");
            }
            if (this.TxtLocPinX)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.TxtLocPinX, "TxtLocPinX  ");
            }
            if (this.TxtLocPinY)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.TxtLocPinY, "TxtLocPinY ");
            }
            if (this.TxtPinX)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.TxtPinX, "TxtPinX ");
            }
            if (this.TxtPinY)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.TxtPinY, "TxtPinY ");
            }
            if (this.TxtWidth)
            {
                query.AddColumn(VA.ShapeSheet.SRCConstants.TxtWidth, "TxtWidth ");
            }


            foreach (var c in query.Columns)
            {
                this.WriteVerboseEx("Column {0} {1}", c.Ordinal,c.Name);
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