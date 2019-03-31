using VisioAutomation.ShapeSheet;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioControl
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioControl)]
    public class GetVisioControl : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter GetCells;

        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.Models.TargetShapes(this.Shapes);
            var dic_shape_to_listofcontrolscells = this.Client.Control.GetControls(targetshapes, CellValueType.Formula);

            if (this.GetCells)
            {
                this.WriteObject(dic_shape_to_listofcontrolscells);
                return;
            }

            foreach (var shape_listofcontrolcells_pair in dic_shape_to_listofcontrolscells)
            {
                var shape = shape_listofcontrolcells_pair.Key;
                var listof_controllcells = shape_listofcontrolcells_pair.Value;
                int shapeid = shape.ID;

                foreach (var controlcells in listof_controllcells)
                {
                    var cp = new VisioPowerShell.Models.Control();

                    cp.ShapeID = shapeid;

                    cp.CanGlue = controlcells.CanGlue.Value;
                    cp.Tip = controlcells.Tip.Value;
                    cp.X = controlcells.X.Value;
                    cp.Y = controlcells.Y.Value;
                    cp.XBehavior = controlcells.XBehavior.Value;
                    cp.YBehavior = controlcells.YBehavior.Value;
                    cp.XDynamics = controlcells.XDynamics.Value;
                    cp.YDynamics = controlcells.YDynamics.Value;

                    this.WriteObject(cp);
                }
            }
        }
    }
}