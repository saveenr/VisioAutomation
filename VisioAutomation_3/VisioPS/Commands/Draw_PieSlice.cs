using VAS = VisioAutomation.Scripting;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Draw", "PieSlice")]
    public class Draw_PieSlice : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public double X0 { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public double Y0 { get; set; }

        [SMA.Parameter(Position = 2, Mandatory = true)]
        public double Radius { get; set; }

        [SMA.Parameter(Position = 3, Mandatory = true)]
        public double StartAngle { get; set; }

        [SMA.Parameter(Position = 4, Mandatory = true)]
        public double EndAngle { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            var shape = scriptingsession.Draw.DrawPieSlice(new VA.Drawing.Point(this.X0, this.Y0), this.Radius, this.StartAngle, this.EndAngle);

            this.WriteObject(shape);
        }
    }
}