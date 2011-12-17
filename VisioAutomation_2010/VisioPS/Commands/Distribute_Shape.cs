using VA= VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Distribute", "Shape")]
    public class Distribute_Shape : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public VA.Drawing.Axis Axis { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)] public double Distance = -1.0;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            if (this.Distance < 0)
            {
                scriptingsession.Layout.Distribute(this.Axis);
            }
            else
            {
                scriptingsession.Layout.Distribute(this.Axis, this.Distance);
            }
        }
    }
}