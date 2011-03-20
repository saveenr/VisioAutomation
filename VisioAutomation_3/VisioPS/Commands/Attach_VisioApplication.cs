using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Attach", "VisioApplication")]
    public class Attach_VisioApplication : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var app = VAS.Session.AttachToRunningApplication();
            Globals.Application = app;
            this.WriteObject(app);
        }
    }
}