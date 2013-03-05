using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using System.Linq;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioShapeProperty")]
    public class Set_VisioShapeProperty: VisioPSCmdlet
    {
        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string Width { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string Height { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string PinX { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string PinY { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string LocPinX { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string LocPinY { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string FillPattern { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string FillForegroundColor { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string FillForegroundtransparency { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string FillBackgroundColor { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string FillBackgroundtransparency { get; set; }
        
        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string LinePattern { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string LineWeight{ get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string LineColor { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string LineCap { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string CharCase { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string CharColor { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]

        public string CharFont { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string CharFontScale { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string CharLetterspace { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string CharSize { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string CharStyle { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string CharColorTransparency { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string BeginX { get; set; }
        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string BeginY{ get; set; }
        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string EndX{ get; set; }
        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string EndY{ get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string Rounding { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string ShadowBackground { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string ShadowBackgroundTransparency { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string ShadowForeground { get; set; }
        
        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string ShadowForegroundTransparency { get; set; }


        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string ShadowObliqueAngle { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string ShadowOffsetX { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string ShadowOffsetY { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string ShadowPattern { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string ShadowScalefactor { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string ShadowType { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string SelectMode { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular;


        [SMA.Parameter(Mandatory = false)]
        public IList<IVisio.Shape> Shapes;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            var update = new VisioAutomation.ShapeSheet.Update();
            update.BlastGuards = this.BlastGuards;
            update.TestCircular= this.TestCircular;

            var target_shapes = this.Shapes ?? scriptingsession.Selection.GetShapes();

            foreach (var shape in target_shapes)
            {
                var id = shape.ID16;
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.Width, this.Width);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.Height, this.Height);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PinX, this.PinX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PinY, this.PinY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.FillForegnd, this.FillForegroundColor);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.FillPattern, this.FillPattern);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.FillBkgnd, this.FillBackgroundColor);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineColor, this.LineColor);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LinePattern, this.LinePattern);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineWeight, this.LineWeight);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.FillForegndTrans, this.FillForegroundtransparency);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.FillBkgndTrans, this.FillBackgroundtransparency);

                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.Char_Case, this.CharCase);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.Char_Color, this.CharColor);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.Char_Font, this.CharFont);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.Char_FontScale, this.CharFontScale);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.Char_Letterspace, this.CharLetterspace);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.Char_Size, this.CharSize);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.Char_Style, this.CharStyle);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.Char_ColorTrans, this.CharColorTransparency);

                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.BeginX, this.BeginX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.EndX, this.EndX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.BeginY, this.BeginY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.EndY, this.EndY);

                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineCap, this.LineCap);

                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.Rounding, this.Rounding);

                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwBkgnd, this.ShadowBackground);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwBkgndTrans, this.ShadowBackgroundTransparency);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwForegnd, this.ShadowForeground);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwForegndTrans, this.ShadowForegroundTransparency);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwObliqueAngle, this.ShadowObliqueAngle);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwOffsetX, this.ShadowOffsetX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwOffsetY, this.ShadowOffsetY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwPattern, this.ShadowPattern);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwScaleFactor, this.ShadowScalefactor);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwType, this.ShadowType);

                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LocPinX, this.LocPinX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LocPinY, this.LocPinY);

                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.SelectMode, this.SelectMode);

            }

            var page = scriptingsession.Page.Get();
            update.Execute(page);
        }
    }
}