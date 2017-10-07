using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Lock, VisioPowerShell.Commands.Nouns.VisioShape)]
    public class LockVisioShape : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

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

        protected override void ProcessRecord()
        {
            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);

            var lockcells = new VisioAutomation.Shapes.LockCells();

            
            lockcells.Aspect =  this.Aspect ? "1": null;
            lockcells.Begin = this.Begin ? "1" : null;
            lockcells.CalcWH = this.CalcWH ? "1" : null;
            lockcells.Crop = this.Crop ? "1" : null;
            lockcells.CustProp = this.CustProp ? "1" : null;
            lockcells.Delete = this.Delete ? "1" : null;
            lockcells.End = this.End ? "1" : null;
            lockcells.Format = this.Format ? "1" : null;
            lockcells.FromGroupFormat =  this.FromGroupFormat ? "1" : null;
            lockcells.Group = this.Group ? "1" : null;
            lockcells.Height = this.Height ? "1" : null;
            lockcells.MoveX = this.MoveX ? "1" : null;
            lockcells.MoveY = this.MoveY ? "1" : null;
            lockcells.Rotate = this.Rotate ? "1" : null;
            lockcells.Select = this.Select ? "1" : null;
            lockcells.TextEdit = this.TextEdit ? "1" : null;
            lockcells.ThemeColors = this.ThemeColors ? "1" : null;
            lockcells.ThemeEffects = this.ThemeEffects ? "1" : null;
            lockcells.VertexEdit = this.VertexEdit ? "1" : null;
            lockcells.Width = this.Width ? "1" : null;

            this.Client.Arrange.SetLock(targets,lockcells);
        }
    }
}
