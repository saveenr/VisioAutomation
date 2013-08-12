using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioPieSlices")]
    public class New_VisioPieSlices : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public double X0 { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public double Y0 { get; set; }

        [SMA.Parameter(Position = 2, Mandatory = true)]
        public double Radius { get; set; }

        [SMA.Parameter(Position = 3, Mandatory = true)] public double[] Values;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var center = new VA.Drawing.Point(this.X0, this.Y0);
            var shapes = scriptingsession.Draw.PieSlices(center, this.Radius, this.Values);
            this.WriteObject(shapes, false);
        }
    }
}