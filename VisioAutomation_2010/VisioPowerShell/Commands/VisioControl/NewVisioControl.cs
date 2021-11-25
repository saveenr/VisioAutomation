


namespace VisioPowerShell.Commands.VisioControl
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.VisioControl)]
    public class NewVisioControl : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public string XDynamics { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string YDynamics { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string XBehavior { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string YBehavior { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string X { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string Y { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public bool CanGlue = false;

        [SMA.Parameter(Mandatory = false)]
        public string Tip { get; set; }

        // CONTEXT:SHAPES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape;

        protected override void ProcessRecord()
        {
            var ctrl = new VA.Shapes.ControlCells();
                ctrl.XDynamics = this.XDynamics;
                ctrl.YDynamics = this.YDynamics;
                ctrl.XBehavior = this.XBehavior;
                ctrl.YBehavior = this.YBehavior;
                ctrl.X = this.X;
                ctrl.Y = this.Y;
                ctrl.CanGlue = this.CanGlue;
                ctrl.Tip = this.Tip;

            var targetshapes = new VisioScripting.TargetShapes(this.Shape);
            this.Client.Control.AddControlToShapes(targetshapes, ctrl);
        }
    }
}