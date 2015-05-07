using System.Linq;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.New, "VisioPolyLine")]
    public class New_VisioPolyLine : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = true)]
        public double[] Doubles { get; set; }

        protected override void ProcessRecord()
        {
            var points = VisioAutomation.Drawing.Point.FromDoubles(this.Doubles).ToList();
            var shape = this.client.Draw.PolyLine(points);
            this.WriteObject(shape);
        }
    }
}