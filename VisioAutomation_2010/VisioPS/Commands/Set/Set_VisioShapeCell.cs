using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using System.Linq;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioShapeCell")]
    public class Set_VisioShapeCell: VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)] public string Width { get; set; }
        [SMA.Parameter(Mandatory = false)] public string Height { get; set; }
        [SMA.Parameter(Mandatory = false)] public string PinX { get; set; }
        [SMA.Parameter(Mandatory = false)] public string PinY { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LocPinX { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LocPinY { get; set; }
        [SMA.Parameter(Mandatory = false)] public string Angle { get; set; }

        [SMA.Parameter(Mandatory = false)] public string FillPattern { get; set; }
        [SMA.Parameter(Mandatory = false)] public string FillForegnd { get; set; }
        [SMA.Parameter(Mandatory = false)] public string FillForegndTrans { get; set; }
        [SMA.Parameter(Mandatory = false)] public string FillBkgnd { get; set; }
        [SMA.Parameter(Mandatory = false)] public string FillBkgndTrans { get; set; }
        
        [SMA.Parameter(Mandatory = false)] public string LinePattern { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LineWeight{ get; set; }
        [SMA.Parameter(Mandatory = false)] public string LineColor { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LineCap { get; set; }
        [SMA.Parameter(Mandatory = false)] public string Rounding { get; set; }

        [SMA.Parameter(Mandatory = false)] public string CharCase { get; set; }
        [SMA.Parameter(Mandatory = false)] public string CharColor { get; set; }
        [SMA.Parameter(Mandatory = false)] public string CharFont { get; set; }
        [SMA.Parameter(Mandatory = false)] public string CharFontScale { get; set; }
        [SMA.Parameter(Mandatory = false)] public string CharLetterspace { get; set; }
        [SMA.Parameter(Mandatory = false)] public string CharSize { get; set; }
        [SMA.Parameter(Mandatory = false)] public string CharStyle { get; set; }
        [SMA.Parameter(Mandatory = false)] public string CharColorTransparency { get; set; }

        [SMA.Parameter(Mandatory = false)] public string BeginX { get; set; }
        [SMA.Parameter(Mandatory = false)] public string BeginY{ get; set; }
        [SMA.Parameter(Mandatory = false)] public string EndX{ get; set; }
        [SMA.Parameter(Mandatory = false)] public string EndY{ get; set; }


        [SMA.Parameter(Mandatory = false)] public string ShdwBkgnd { get; set; }
        [SMA.Parameter(Mandatory = false)] public string ShdwBkgndTrans { get; set; }
        [SMA.Parameter(Mandatory = false)] public string ShdwForegnd { get; set; }
        
        [SMA.Parameter(Mandatory = false)] public string ShdwForegndTrans { get; set; }
        [SMA.Parameter(Mandatory = false)] public string ShdwObliqueAngle { get; set; }
        [SMA.Parameter(Mandatory = false)] public string ShdwOffsetX { get; set; }
        [SMA.Parameter(Mandatory = false)] public string ShdwOffsetY { get; set; }
        [SMA.Parameter(Mandatory = false)] public string ShdwPattern { get; set; }
        [SMA.Parameter(Mandatory = false)] public string ShdwScalefactor { get; set; }
        [SMA.Parameter(Mandatory = false)] public string ShdwType { get; set; }
        [SMA.Parameter(Mandatory = false)] public string SelectMode { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards { get; set; }
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular { get; set; }


        [SMA.Parameter(Mandatory = false)]
        public string LockAspect { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockBegin { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockCalcWH { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockCrop { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockCustProp { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockDelete { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockEnd { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockFormat { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockFromGroupFormat { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockGroup { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockHeight { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockMoveX { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockMoveY { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockRotate { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockSelect { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockTextEdit { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockThemeColors { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockThemeEffects { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockVtxEdit { get; set; }
        [SMA.Parameter(Mandatory = false)] public string LockWidth { get; set; }

        [SMA.Parameter(Mandatory = false)] public string TxtAngle { get; set; }
        [SMA.Parameter(Mandatory = false)] public string TxtHeight { get; set; }
        [SMA.Parameter(Mandatory = false)] public string TxtLocPinX  { get; set; }
        [SMA.Parameter(Mandatory = false)] public string TxtLocPinY { get; set; }
        [SMA.Parameter(Mandatory = false)] public string TxtPinX { get; set; }
        [SMA.Parameter(Mandatory = false)] public string TxtPinY { get; set; }
        [SMA.Parameter(Mandatory = false)] public string TxtWidth { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes { get; set; }

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
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.Angle, this.Angle);
 
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.FillForegnd, this.FillForegnd);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.FillPattern, this.FillPattern);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.FillBkgnd, this.FillBkgnd);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineColor, this.LineColor);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LinePattern, this.LinePattern);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineWeight, this.LineWeight);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.FillForegndTrans, this.FillForegndTrans);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.FillBkgndTrans, this.FillBkgndTrans);

                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.CharCase, this.CharCase);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.CharColor, this.CharColor);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.CharFont, this.CharFont);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.CharFontScale, this.CharFontScale);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.CharLetterspace, this.CharLetterspace);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.CharSize, this.CharSize);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.CharStyle, this.CharStyle);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.CharColorTrans, this.CharColorTransparency);

                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.BeginX, this.BeginX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.EndX, this.EndX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.BeginY, this.BeginY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.EndY, this.EndY);

                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineCap, this.LineCap);

                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.Rounding, this.Rounding);

                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwBkgnd, this.ShdwBkgnd);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwBkgndTrans, this.ShdwBkgndTrans);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwForegnd, this.ShdwForegnd);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwForegndTrans, this.ShdwForegndTrans);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwObliqueAngle, this.ShdwObliqueAngle);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwOffsetX, this.ShdwOffsetX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwOffsetY, this.ShdwOffsetY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwPattern, this.ShdwPattern);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwScaleFactor, this.ShdwScalefactor);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.ShdwType, this.ShdwType);

                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LocPinX, this.LocPinX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LocPinY, this.LocPinY);

                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.SelectMode, this.SelectMode);
                
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockAspect , this.LockAspect);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockBegin, this.LockBegin);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockCalcWH, this.LockCalcWH);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockCrop, this.LockCrop);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockCustProp, this.LockCustProp);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockDelete, this.LockDelete);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockEnd, this.LockEnd);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockFormat, this.LockFormat);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockFromGroupFormat, this.LockFromGroupFormat);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockGroup, this.LockGroup);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockHeight, this.LockHeight);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockMoveX, this.LockMoveX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockMoveY, this.LockMoveY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockRotate, this.LockRotate);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockSelect, this.LockSelect);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockTextEdit, this.LockTextEdit);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockThemeColors, this.LockThemeColors);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockThemeEffects, this.LockThemeEffects);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockVtxEdit, this.LockVtxEdit);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LockWidth, this.LockWidth);

                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.TxtAngle, this.TxtAngle);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.TxtHeight, this.TxtHeight);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.TxtLocPinX, this.TxtHeight);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.TxtLocPinY, this.TxtLocPinY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.TxtPinX, this.TxtPinX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.TxtPinY, this.TxtPinY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.TxtWidth, this.TxtWidth);
            }

            var page = scriptingsession.Page.Get();

            this.WriteVerboseEx("BlastGuards: {0}", this.BlastGuards);
            this.WriteVerboseEx("TestCircular: {0}", this.TestCircular);
            this.WriteVerboseEx("Number of Shapes : {0}", target_shapes.Count);
            this.WriteVerboseEx("Number of Total Updates: {0}", update.Count());
            this.WriteVerboseEx("Number of Updates per Shape: {0}", update.Count() / target_shapes.Count());

            using (var undoscope = new VA.Application.UndoScope(this.ScriptingSession.VisioApplication, "SetShapeCells"))
            {
                this.WriteVerboseEx("Start Update");
                update.Execute(page);
                this.WriteVerboseEx("End Update");
            }
        }
    }
}