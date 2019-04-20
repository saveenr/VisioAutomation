using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Unlock, Nouns.VisioShape)]
    public class UnlockVisioShape : VisioCmdlet
    {
        public SMA.SwitchParameter Aspect;
        public SMA.SwitchParameter Begin;
        public SMA.SwitchParameter CalcWH;
        public SMA.SwitchParameter Crop;
        public SMA.SwitchParameter CustProp;
        public SMA.SwitchParameter Delete;
        public SMA.SwitchParameter End;
        public SMA.SwitchParameter Format;
        public SMA.SwitchParameter FromGroupFormat;
        public SMA.SwitchParameter Group;
        public SMA.SwitchParameter Height;
        public SMA.SwitchParameter MoveX;
        public SMA.SwitchParameter MoveY;
        public SMA.SwitchParameter Rotate;
        public SMA.SwitchParameter Select;
        public SMA.SwitchParameter TextEdit;
        public SMA.SwitchParameter ThemeColors;
        public SMA.SwitchParameter ThemeEffects;
        public SMA.SwitchParameter VertexEdit;
        public SMA.SwitchParameter Width;

        // CONTEXT:SHAPES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shapes);

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
