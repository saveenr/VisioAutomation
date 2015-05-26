using System.Management.Automation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Get
{
    [Cmdlet(VerbsCommon.Get, "VisioConnectionPoint")]
    public class Get_VisioConnectionPoint : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [Parameter(Mandatory = false)]
        public SwitchParameter GetCells;

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

                foreach (var point_cells in points)
                {
                    var cp = new Model.ConnectionPointValues(shapeid, point_cells);
                    this.WriteObject(cp);
                }
            }
        }

    }
}