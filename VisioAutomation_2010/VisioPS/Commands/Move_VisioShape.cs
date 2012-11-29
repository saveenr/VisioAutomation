using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Move", "VisioShape")]
    public class Move_VisioShape : VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public double Left { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public double Right { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public double Up { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public double Down { get; set; }


        protected override void ProcessRecord()
        {
            double h = Right - Left;
            double v = Up - Down;

            var scriptingsession = this.ScriptingSession;
            scriptingsession.Layout.Nudge(h, v);
        }
    }
}