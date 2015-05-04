using System.Linq;
using VisioAutomation.Drawing;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.New, "VisioBezier")]
    public class New_VisioBezier : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = true)]
        public double[] Doubles { get; set; }

        protected override void ProcessRecord()
        {
            var points = Point.FromDoubles(this.Doubles).ToList();
            var shape = this.client.Draw.Bezier(points);
            this.WriteObject(shape);
        }
    }
}