using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Get, VisioPowerShell.Commands.Nouns.VisioControl)]
    public class GetVisioControl : VisioCmdlet
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

                    cp.CanGlue = point.CanGlue.ValueF;
                    cp.Tip = point.Tip.ValueF;
                    cp.X = point.X.ValueF;
                    cp.Y = point.Y.ValueF;
                    cp.XBehavior = point.XBehavior.ValueF;
                    cp.YBehavior = point.YBehavior.ValueF;
                    cp.XDynamics = point.XDynamics.ValueF;
                    cp.YDynamics = point.YDynamics.ValueF;

                    this.WriteObject(cp);
                }
            }
        }
    }
}