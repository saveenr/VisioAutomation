using System.Collections.Generic;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsLifecycle.Invoke, "VisioNudgeShape")]
    public class Invoke_VisioNudgeShape : VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public double Left { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public double Right { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public double Up { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public double Down { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            double h = Right - Left;
            double v = Up - Down;

            var scriptingsession = this.ScriptingSession;
            scriptingsession.Layout.Nudge(this.Shapes, h, v);
        }
    }
}