using System.Linq;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioBezier")]
    public class New_VisioBezier : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public double[] Doubles { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var points = VA.Drawing.Point.FromDoubles(this.Doubles).ToList();
            var shape = scriptingsession.Draw.Bezier(points);
            this.WriteObject(shape);
        }
    }
}