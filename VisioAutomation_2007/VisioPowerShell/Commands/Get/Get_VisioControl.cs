using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioControl")]
    public class Get_VisioControl : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter GetCells;

        protected override void ProcessRecord()
        {
            var dic = this.client.Control.Get(this.Shapes);

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
                    var cp = new ControlFormulas();

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

    public class ControlFormulas
    {
        public int ShapeID;
        public string CanGlue;
        public string Tip;
        public string X;
        public string Y;
        public string XBehavior;
        public string YBehavior;
        public string XDynamics;
        public string YDynamics;
    }
}