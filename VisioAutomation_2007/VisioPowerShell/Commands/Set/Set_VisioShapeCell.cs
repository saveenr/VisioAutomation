using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using System.Linq;
using VA = VisioAutomation;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioShapeCell")]
    public class Set_VisioShapeCell : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false, Position = 0)]
        public System.Collections.Hashtable Hashtable { get; set; }

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
            var update = new VisioAutomation.ShapeSheet.Update();
            update.BlastGuards = this.BlastGuards;
            update.TestCircular = this.TestCircular;

            var valuemap = new CellValueMap(Get_VisioShapeCell.GetShapeCellDictionary());
            valuemap.UpdateValueMap(this.Hashtable);
            valuemap.SetIf("Angle", this.Angle);
            valuemap.SetIf("BeginArrow", this.BeginArrow);
            valuemap.SetIf("BeginArrowSize", this.BeginArrowSize);
            valuemap.SetIf("BeginX", this.BeginX);
            valuemap.SetIf("BeginY", this.BeginY);
            valuemap.SetIf("CharCase", this.CharCase);
            valuemap.SetIf("CharColor", this.CharColor);
            valuemap.SetIf("CharColorTransparency", this.CharColorTransparency);
            valuemap.SetIf("CharFont", this.CharFont);
            valuemap.SetIf("CharFontScale", this.CharFontScale);
            valuemap.SetIf("CharLetterspace", this.CharLetterspace);
            valuemap.SetIf("CharSize", this.CharSize);
            valuemap.SetIf("CharStyle", this.CharStyle);
            valuemap.SetIf("EndArrow", this.EndArrow);
            valuemap.SetIf("EndArrowSize", this.EndArrowSize);
            valuemap.SetIf("EndX", this.EndX);
            valuemap.SetIf("EndY", this.EndY);
            valuemap.SetIf("FillBkgnd", this.FillBkgnd);
            valuemap.SetIf("FillBkgndTrans", this.FillBkgndTrans);
            valuemap.SetIf("FillForegnd", this.FillForegnd);
            valuemap.SetIf("FillForegndTrans", this.FillForegndTrans);
            valuemap.SetIf("FillPattern", this.FillPattern);
            valuemap.SetIf("Height", this.Height);
            valuemap.SetIf("HideText", this.HideText);
            valuemap.SetIf("LineCap", this.LineCap);
            valuemap.SetIf("LineColor", this.LineColor);
            valuemap.SetIf("LinePattern", this.LinePattern);
            valuemap.SetIf("LineWeight", this.LineWeight);
            valuemap.SetIf("LockAspect", this.LockAspect);
            valuemap.SetIf("LockBegin", this.LockBegin);
            valuemap.SetIf("LockCalcWH", this.LockCalcWH);
            valuemap.SetIf("LockCrop", this.LockCrop);
            valuemap.SetIf("LockCustProp", this.LockCustProp);
            valuemap.SetIf("LockDelete", this.LockDelete);
            valuemap.SetIf("LockEnd", this.LockEnd);
            valuemap.SetIf("LockFormat", this.LockFormat);
            valuemap.SetIf("LockFromGroupFormat", this.LockFromGroupFormat);
            valuemap.SetIf("LockGroup", this.LockGroup);
            valuemap.SetIf("LockHeight", this.LockHeight);
            valuemap.SetIf("LockMoveX", this.LockMoveX);
            valuemap.SetIf("LockMoveY", this.LockMoveY);
            valuemap.SetIf("LockRotate", this.LockRotate);
            valuemap.SetIf("LockSelect", this.LockSelect);
            valuemap.SetIf("LockTextEdit", this.LockTextEdit);
            valuemap.SetIf("LockThemeColors", this.LockThemeColors);
            valuemap.SetIf("LockThemeEffects", this.LockThemeEffects);
            valuemap.SetIf("LockVtxEdit", this.LockVtxEdit);
            valuemap.SetIf("LockWidth", this.LockWidth);
            valuemap.SetIf("LocPinX", this.LocPinX);
            valuemap.SetIf("LocPinY", this.LocPinY);
            valuemap.SetIf("PinX", this.PinX);
            valuemap.SetIf("PinY", this.PinY);
            valuemap.SetIf("Rounding", this.Rounding);
            valuemap.SetIf("SelectMode", this.SelectMode);
            valuemap.SetIf("ShdwBkgnd", this.ShdwBkgnd);
            valuemap.SetIf("ShdwBkgndTrans", this.ShdwBkgndTrans);
            valuemap.SetIf("ShdwForegnd", this.ShdwForegnd);
            valuemap.SetIf("ShdwForegndTrans", this.ShdwForegndTrans);
            valuemap.SetIf("ShdwObliqueAngle", this.ShdwObliqueAngle);
            valuemap.SetIf("ShdwOffsetX", this.ShdwOffsetX);
            valuemap.SetIf("ShdwOffsetY", this.ShdwOffsetY);
            valuemap.SetIf("ShdwPattern", this.ShdwPattern);
            valuemap.SetIf("ShdwScalefactor", this.ShdwScalefactor);
            valuemap.SetIf("ShdwType", this.ShdwType);
            valuemap.SetIf("TxtAngle", this.TxtAngle);
            valuemap.SetIf("TxtHeight", this.TxtHeight);
            valuemap.SetIf("TxtHeight", this.TxtHeight);
            valuemap.SetIf("TxtLocPinY", this.TxtLocPinY);
            valuemap.SetIf("TxtPinX", this.TxtPinX);
            valuemap.SetIf("TxtPinY", this.TxtPinY);
            valuemap.SetIf("TxtWidth", this.TxtWidth);
            valuemap.SetIf("Width", this.Width);

            var target_shapes = this.Shapes ?? this.client.Selection.GetShapes();

            foreach (var shape in target_shapes)
            {
                var id = shape.ID16;

                foreach (var cellname in valuemap.CellNames)
                {
                    string cell_value = valuemap[cellname];
                    var cell_src = valuemap.GetSRC(cellname);
                    update.SetFormulaIgnoreNull(id,cell_src, cell_value);
                }
            }

            var surface = this.client.Draw.GetDrawingSurfaceSafe();

            this.WriteVerbose("BlastGuards: {0}", this.BlastGuards);
            this.WriteVerbose("TestCircular: {0}", this.TestCircular);
            this.WriteVerbose("Number of Shapes : {0}", target_shapes.Count);
            this.WriteVerbose("Number of Total Updates: {0}", update.Count());

            using (var undoscope = new VA.Application.UndoScope(this.client.VisioApplication, "SetShapeCells"))
            {
                this.WriteVerbose("Start Update");
                update.Execute(surface);
                this.WriteVerbose("End Update");
            }
        }
    }
}