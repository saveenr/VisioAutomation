using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Unlock, Nouns.VisioShape)]
    public class UnlockVisioShape : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter Aspect;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter Begin;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter CalcWH;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter Crop;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter CustProp;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter Delete;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter End;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter Format;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter FromGroupFormat;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter Group;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter Height;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter MoveX;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter MoveY;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter Rotate;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter Select;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter TextEdit;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter ThemeColors;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter ThemeEffects;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter VertexEdit;
        [SMA.Parameter(Mandatory = false)] public SMA.SwitchParameter Width;

        // CONTEXT:SHAPES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape;

        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shape);

            var lockcells = new VisioAutomation.Shapes.LockCells();

            lockcells.Aspect =  this.Aspect ? "0" : null;
            lockcells.Begin = this.Begin ? "0" : null;
            lockcells.CalcWH = this.CalcWH ? "0" : null;
            lockcells.Crop = this.Crop ? "0" : null;
            lockcells.CustProp = this.CustProp ? "0" : null;
            lockcells.Delete = this.Delete ? "0" : null;
            lockcells.End = this.End ? "0" : null;
            lockcells.Format = this.Format ? "0" : null;
            lockcells.FromGroupFormat = this.FromGroupFormat ? "0" : null;
            lockcells.Group = this.Group ? "0" : null;
            lockcells.Height = this.Height ? "0" : null;
            lockcells.MoveX = this.MoveX ? "0" : null;
            lockcells.MoveY = this.MoveY ? "0" : null;
            lockcells.Rotate = this.Rotate ? "0" : null;
            lockcells.Select = this.Select ? "0" : null;
            lockcells.TextEdit = this.TextEdit ? "0" : null;
            lockcells.ThemeColors = this.ThemeColors ? "0" : null;
            lockcells.ThemeEffects = this.ThemeEffects ? "0" : null;
            lockcells.VertexEdit = this.VertexEdit ? "0" : null;
            lockcells.Width = this.Width ? "0" : null;

            this.Client.Lock.SetLockCells(targetshapes, lockcells);
        }
    }
}
