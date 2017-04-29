using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Get, VisioPowerShell.Commands.Nouns.VisioControl)]
    public class Get_VisioControl : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [Parameter(Mandatory = false)]
        public SwitchParameter GetCells;

        protected override void ProcessRecord()
        {
            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);
            var dic = this.Client.Control.Get(targets);

            if (this.GetCells)
            {
                this.WriteObject(dic);
                return;
            }

            foreach (var shape_points in dic)
            {
                var shape = shape_points.Key;
                var points = shape_points.Value;
                int shapeid = shape.ID;

                foreach (var point in points)
                {
                    var cp = new VisioPowerShell.Models.Control();

                    cp.ShapeID = shapeid;

                    cp.CanGlue = point.CanGlue.Formula.Value;
                    cp.Tip = point.Tip.Formula.Value;
                    cp.X = point.X.Formula.Value;
                    cp.Y = point.Y.Formula.Value;
                    cp.XBehavior = point.XBehavior.Formula.Value;
                    cp.YBehavior = point.YBehavior.Formula.Value;
                    cp.XDynamics = point.XDynamics.Formula.Value;
                    cp.YDynamics = point.YDynamics.Formula.Value;

                    this.WriteObject(cp);
                }
            }
        }
    }
}