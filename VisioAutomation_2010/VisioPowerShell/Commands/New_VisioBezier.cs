using System.Linq;
using System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioBezier)]
    public class New_VisioBezier : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public double[] Doubles { get; set; }

        protected override void ProcessRecord()
        {
            var points = VA.Drawing.Point.FromDoubles(this.Doubles).ToList();
            var shape = this.Client.Draw.Bezier(points);
            this.WriteObject(shape);
        }
    }
}