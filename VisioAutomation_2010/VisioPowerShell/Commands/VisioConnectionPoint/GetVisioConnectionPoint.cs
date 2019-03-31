using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioConnectionPoint
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioConnectionPoint)]
    public class GetVisioConnectionPoint : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter GetCells;

        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.Models.TargetShapes(this.Shapes);

            var dic = this.Client.ConnectionPoint.GetConnectionPoints(targetshapes);

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

                foreach (var point_cells in points)
                {
                    var cp = new Models.ConnectionPoint(shapeid, point_cells);
                    this.WriteObject(cp);
                }
            }
        }

    }
}