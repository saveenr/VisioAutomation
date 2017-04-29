using System.Management.Automation;
using VisioAutomation.Shapes;
using VisioAutomation.Utilities;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioControl)]
    public class NewVisioControl : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public string XDynamics { get; set; }

        [Parameter(Mandatory = false)]
        public string YDynamics { get; set; }

        [Parameter(Mandatory = false)]
        public string XBehavior { get; set; }

        [Parameter(Mandatory = false)]
        public string YBehavior { get; set; }

        [Parameter(Mandatory = false)]
        public string X { get; set; }

        [Parameter(Mandatory = false)]
        public string Y { get; set; }

        [Parameter(Mandatory = false)]
        public bool CanGlue = false;

        [Parameter(Mandatory = false)]
        public string Tip { get; set; }

        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var ctrl = new ControlCells();
                ctrl.XDynamics = this.XDynamics;
                ctrl.YDynamics = this.YDynamics;
                ctrl.XBehavior = this.XBehavior;
                ctrl.YBehavior = this.YBehavior;
                ctrl.X = this.X;
                ctrl.Y = this.Y;
                ctrl.CanGlue = Convert.BoolToFormula(this.CanGlue);
                ctrl.Tip = this.Tip;

            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);

            this.Client.Control.Add(targets, ctrl);
        }
    }
}