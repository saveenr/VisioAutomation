using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioConnectionPoint")]
    public class Get_VisioConnectionPoint : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]public IVisio.Shape[] Shapes;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter GetCells;

        protected override void ProcessRecord()
        {
            var dic = this.client.ConnectionPoint.Get(this.Shapes);

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
                    var cp = new ConnectionPointValues();

                    cp.ShapeID = shapeid;

                    cp.Type = point.Type.Formula.Value;
                    cp.X = point.X.Formula.Value;
                    cp.Y = point.Y.Formula.Value;
                    cp.DirX = point.DirX.Formula.Value;
                    cp.DirY = point.DirY.Formula.Value;

                    this.WriteObject(cp);
                }
            }
        }
    }
}