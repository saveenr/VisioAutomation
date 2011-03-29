using System.Linq;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Draw", "Bezier")]
    public class Draw_Bezier : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public double[] Doubles { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var points = VA.Drawing.DrawingUtil.DoublesToPoints(this.Doubles).ToList();
            var shape = scriptingsession.Draw.DrawBezier(points);
            this.WriteObject(shape);
        }
    }
}