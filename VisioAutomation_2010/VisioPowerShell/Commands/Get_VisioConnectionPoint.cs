using System.Management.Automation;
using VisioScripting.Models;
using VisioPowerShell.Models;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Get, VisioPowerShell.Nouns.VisioConnectionPoint)]
    public class Get_VisioConnectionPoint : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        [Parameter(Mandatory = false)]
        public SwitchParameter GetCells;

        protected override void ProcessRecord()
        {
            var targets = new TargetShapes(this.Shapes);

            var dic = this.Client.ConnectionPoint.Get(targets);

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
                    var cp = new ConnectionPoint(shapeid, point_cells);
                    this.WriteObject(cp);
                }
            }
        }

    }
}