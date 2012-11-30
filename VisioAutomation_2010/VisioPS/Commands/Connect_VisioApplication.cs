using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommunications.Connect, "VisioApplication")]
    public class Connect_VisioApplication : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var app = VA.Application.ApplicationHelper.FindRunningApplication();
            Globals.Application = app;
            this.WriteObject(app);
        }
    }
}