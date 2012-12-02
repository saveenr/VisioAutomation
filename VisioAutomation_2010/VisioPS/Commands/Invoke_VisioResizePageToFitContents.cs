using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsLifecycle.Invoke, "VisioResizePageToFitContents")]
    public class Invoke_VisioResizePageToFitContents : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public double BorderWidth { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public double BorderHeight { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var bordersize = new VA.Drawing.Size(BorderWidth, BorderWidth);
            scriptingsession.Page.ResizeToFitContents(bordersize, true);
        }
    }
}