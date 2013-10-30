using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using System.Linq;
using VA = VisioAutomation;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioShapeCell")]
    public class Set_VisioShapeCell : VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public string Width { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string Height { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string PinX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string PinY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LocPinX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LocPinY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string Angle { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string FillPattern { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string FillForegnd { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string FillForegndTrans { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string FillBkgnd { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string FillBkgndTrans { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LinePattern { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LineWeight { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LineColor { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LineCap { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string Rounding { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string CharCase { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string CharColor { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string CharFont { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string CharFontScale { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string CharLetterspace { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string CharSize { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string CharStyle { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string CharColorTransparency { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string BeginArrow { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string BeginArrowSize { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string EndArrow { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string EndArrowSize { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string BeginX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string BeginY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string EndX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string EndY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string ShdwBkgnd { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string ShdwBkgndTrans { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string ShdwForegnd { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string ShdwForegndTrans { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string ShdwObliqueAngle { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string ShdwOffsetX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string ShdwOffsetY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string ShdwPattern { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string ShdwScalefactor { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string ShdwType { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string SelectMode { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockAspect { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockBegin { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockCalcWH { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockCrop { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockCustProp { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockDelete { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockEnd { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockFormat { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockFromGroupFormat { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockGroup { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockHeight { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockMoveX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockMoveY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockRotate { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockSelect { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockTextEdit { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockThemeColors { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockThemeEffects { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockVtxEdit { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string LockWidth { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string TxtAngle { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string TxtHeight { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string TxtLocPinX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string TxtLocPinY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string TxtPinX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string TxtPinY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string TxtWidth { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string HideText { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            var update = new VisioAutomation.ShapeSheet.Update();
            update.BlastGuards = this.BlastGuards;
            update.TestCircular = this.TestCircular;

            var target_shapes = this.Shapes ?? scriptingsession.Selection.GetShapes();

            foreach (var shape in target_shapes)
            {
                var id = shape.ID16;
                update.SetFormulaIgnoreNull(id, SRCCON.Width, this.Width);
                update.SetFormulaIgnoreNull(id, SRCCON.Height, this.Height);
                update.SetFormulaIgnoreNull(id, SRCCON.PinX, this.PinX);
                update.SetFormulaIgnoreNull(id, SRCCON.PinY, this.PinY);
                update.SetFormulaIgnoreNull(id, SRCCON.Angle, this.Angle);

                update.SetFormulaIgnoreNull(id, SRCCON.FillForegnd, this.FillForegnd);
                update.SetFormulaIgnoreNull(id, SRCCON.FillPattern, this.FillPattern);
                update.SetFormulaIgnoreNull(id, SRCCON.FillBkgnd, this.FillBkgnd);
                update.SetFormulaIgnoreNull(id, SRCCON.LineColor, this.LineColor);
                update.SetFormulaIgnoreNull(id, SRCCON.LinePattern, this.LinePattern);
                update.SetFormulaIgnoreNull(id, SRCCON.LineWeight, this.LineWeight);
                update.SetFormulaIgnoreNull(id, SRCCON.FillForegndTrans, this.FillForegndTrans);
                update.SetFormulaIgnoreNull(id, SRCCON.FillBkgndTrans, this.FillBkgndTrans);
                update.SetFormulaIgnoreNull(id, SRCCON.CharCase, this.CharCase);
                update.SetFormulaIgnoreNull(id, SRCCON.CharColor, this.CharColor);
                update.SetFormulaIgnoreNull(id, SRCCON.CharFont, this.CharFont);
                update.SetFormulaIgnoreNull(id, SRCCON.CharFontScale, this.CharFontScale);
                update.SetFormulaIgnoreNull(id, SRCCON.CharLetterspace, this.CharLetterspace);
                update.SetFormulaIgnoreNull(id, SRCCON.CharSize, this.CharSize);
                update.SetFormulaIgnoreNull(id, SRCCON.CharStyle, this.CharStyle);
                update.SetFormulaIgnoreNull(id, SRCCON.CharColorTrans, this.CharColorTransparency);
                update.SetFormulaIgnoreNull(id, SRCCON.BeginArrow, this.BeginArrow);
                update.SetFormulaIgnoreNull(id, SRCCON.BeginArrowSize, this.BeginArrowSize);
                update.SetFormulaIgnoreNull(id, SRCCON.EndArrow, this.EndArrow);
                update.SetFormulaIgnoreNull(id, SRCCON.EndArrowSize, this.EndArrowSize);

                update.SetFormulaIgnoreNull(id, SRCCON.BeginX, this.BeginX);
                update.SetFormulaIgnoreNull(id, SRCCON.EndX, this.EndX);
                update.SetFormulaIgnoreNull(id, SRCCON.BeginY, this.BeginY);
                update.SetFormulaIgnoreNull(id, SRCCON.EndY, this.EndY);
                update.SetFormulaIgnoreNull(id, SRCCON.LineCap, this.LineCap);
                update.SetFormulaIgnoreNull(id, SRCCON.Rounding, this.Rounding);
                update.SetFormulaIgnoreNull(id, SRCCON.ShdwBkgnd, this.ShdwBkgnd);
                update.SetFormulaIgnoreNull(id, SRCCON.ShdwBkgndTrans, this.ShdwBkgndTrans);
                update.SetFormulaIgnoreNull(id, SRCCON.ShdwForegnd, this.ShdwForegnd);
                update.SetFormulaIgnoreNull(id, SRCCON.ShdwForegndTrans, this.ShdwForegndTrans);
                update.SetFormulaIgnoreNull(id, SRCCON.ShdwObliqueAngle, this.ShdwObliqueAngle);
                update.SetFormulaIgnoreNull(id, SRCCON.ShdwOffsetX, this.ShdwOffsetX);
                update.SetFormulaIgnoreNull(id, SRCCON.ShdwOffsetY, this.ShdwOffsetY);
                update.SetFormulaIgnoreNull(id, SRCCON.ShdwPattern, this.ShdwPattern);
                update.SetFormulaIgnoreNull(id, SRCCON.ShdwScaleFactor, this.ShdwScalefactor);
                update.SetFormulaIgnoreNull(id, SRCCON.ShdwType, this.ShdwType);
                update.SetFormulaIgnoreNull(id, SRCCON.LocPinX, this.LocPinX);
                update.SetFormulaIgnoreNull(id, SRCCON.LocPinY, this.LocPinY);
                update.SetFormulaIgnoreNull(id, SRCCON.SelectMode, this.SelectMode);
                update.SetFormulaIgnoreNull(id, SRCCON.LockAspect, this.LockAspect);
                update.SetFormulaIgnoreNull(id, SRCCON.LockBegin, this.LockBegin);
                update.SetFormulaIgnoreNull(id, SRCCON.LockCalcWH, this.LockCalcWH);
                update.SetFormulaIgnoreNull(id, SRCCON.LockCrop, this.LockCrop);
                update.SetFormulaIgnoreNull(id, SRCCON.LockCustProp, this.LockCustProp);
                update.SetFormulaIgnoreNull(id, SRCCON.LockDelete, this.LockDelete);
                update.SetFormulaIgnoreNull(id, SRCCON.LockEnd, this.LockEnd);
                update.SetFormulaIgnoreNull(id, SRCCON.LockFormat, this.LockFormat);
                update.SetFormulaIgnoreNull(id, SRCCON.LockFromGroupFormat, this.LockFromGroupFormat);
                update.SetFormulaIgnoreNull(id, SRCCON.LockGroup, this.LockGroup);
                update.SetFormulaIgnoreNull(id, SRCCON.LockHeight, this.LockHeight);
                update.SetFormulaIgnoreNull(id, SRCCON.LockMoveX, this.LockMoveX);
                update.SetFormulaIgnoreNull(id, SRCCON.LockMoveY, this.LockMoveY);
                update.SetFormulaIgnoreNull(id, SRCCON.LockRotate, this.LockRotate);
                update.SetFormulaIgnoreNull(id, SRCCON.LockSelect, this.LockSelect);
                update.SetFormulaIgnoreNull(id, SRCCON.LockTextEdit, this.LockTextEdit);
                update.SetFormulaIgnoreNull(id, SRCCON.LockThemeColors, this.LockThemeColors);
                update.SetFormulaIgnoreNull(id, SRCCON.LockThemeEffects, this.LockThemeEffects);
                update.SetFormulaIgnoreNull(id, SRCCON.LockVtxEdit, this.LockVtxEdit);
                update.SetFormulaIgnoreNull(id, SRCCON.LockWidth, this.LockWidth);
                update.SetFormulaIgnoreNull(id, SRCCON.TxtAngle, this.TxtAngle);
                update.SetFormulaIgnoreNull(id, SRCCON.TxtHeight, this.TxtHeight);
                update.SetFormulaIgnoreNull(id, SRCCON.TxtLocPinX, this.TxtHeight);
                update.SetFormulaIgnoreNull(id, SRCCON.TxtLocPinY, this.TxtLocPinY);
                update.SetFormulaIgnoreNull(id, SRCCON.TxtPinX, this.TxtPinX);
                update.SetFormulaIgnoreNull(id, SRCCON.TxtPinY, this.TxtPinY);
                update.SetFormulaIgnoreNull(id, SRCCON.TxtWidth, this.TxtWidth);

                update.SetFormulaIgnoreNull(id, SRCCON.HideText, this.HideText);
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