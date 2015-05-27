using System.Linq;
using System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands.New
{
    [Cmdlet(VerbsCommon.New, "VisioNURBS")]
    public class New_VisioNURBS : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public double[] ControlPoints { get; set; }

        [Parameter(Position = 1, Mandatory = true)]
        public double[] Knots { get; set; }

        [Parameter(Position = 2, Mandatory = true)]
        public double[] Weights { get; set; }

        [Parameter(Position = 3, Mandatory = true)]
        public int Degree { get; set; }
        
        protected override void ProcessRecord()
        {
            var points = VA.Drawing.Point.FromDoubles(this.ControlPoints).ToList();
            var shape = this.client.Draw.NURBSCurve(points, this.Knots, this.Weights, this.Degree);
            this.WriteObject(shape);
        }
    }
}