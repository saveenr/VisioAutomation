using VisioAutomation.Shapes.Controls;
using VA = VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.New, "VisioControl")]
    public class New_VisioControl : VisioCmdlet
    {
        [SMA.ParameterAttribute(Mandatory = false)]
        public string XDynamics { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
        public string YDynamics { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
        public string XBehavior { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
        public string YBehavior { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
        public string X { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
        public string Y { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
        public bool CanGlue = false;

        [SMA.ParameterAttribute(Mandatory = false)]
        public string Tip { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
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
                ctrl.CanGlue = VA.Convert.BoolToFormula(this.CanGlue);
                ctrl.Tip = this.Tip;

                this.client.Control.Add(this.Shapes, ctrl);
        }
    }
}