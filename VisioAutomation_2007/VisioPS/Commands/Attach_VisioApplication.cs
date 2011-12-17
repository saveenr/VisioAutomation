using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Attach", "VisioApplication")]
    public class Attach_VisioApplication : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var app = VA.ApplicationHelper.FindRunningApplication();
            Globals.Application = app;
            this.WriteObject(app);
        }
    }
}