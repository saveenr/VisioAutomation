using VA = VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioControl")]
    public class New_VisioControl : VisioCmdlet
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

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var ctrl = new VA.Shapes.Controls.ControlCells();
                ctrl.XDynamics = this.XDynamics;
                ctrl.YDynamics = this.YDynamics;
                ctrl.XBehavior = this.XBehavior;
                ctrl.YBehavior = this.YBehavior;
                ctrl.X = this.X;
                ctrl.Y = this.Y;
                ctrl.CanGlue = VA.Convert.BoolToFormula(this.CanGlue);
                ctrl.Tip = this.Tip;

                this.client.Control.Add(this.Shapes, ctrl);
        }
    }
}