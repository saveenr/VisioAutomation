using System.Linq;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.New, "VisioNURBS")]
    public class New_VisioNURBS : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = true)]
        public double[] ControlPoints { get; set; }

        [SMA.ParameterAttribute(Position = 1, Mandatory = true)]
        public double[] Knots { get; set; }

        [SMA.ParameterAttribute(Position = 2, Mandatory = true)]
        public double[] Weights { get; set; }

        [SMA.ParameterAttribute(Position = 3, Mandatory = true)]
        public int Degree { get; set; }
        
        protected override void ProcessRecord()
        {
            var points = VA.Drawing.Point.FromDoubles(this.ControlPoints).ToList();
            var shape = this.client.Draw.NURBSCurve(points, this.Knots, this.Weights, this.Degree);
            this.WriteObject(shape);
        }
    }
}