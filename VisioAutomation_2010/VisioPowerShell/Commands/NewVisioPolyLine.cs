using System.Linq;
using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioPolyLine)]
    public class NewVisioPolyLine : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public double[] Doubles { get; set; }

        protected override void ProcessRecord()
        {
            var points = VisioAutomation.Geometry.Point.FromDoubles(this.Doubles).ToList();
            var shape = this.Client.Draw.PolyLine(points);
            this.WriteObject(shape);
        }
    }
}