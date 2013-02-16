using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommunications.Connect, "VisioApplication")]
    public class Connect_VisioApplication : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            if (Globals.Application != null)
            {
                this.WriteWarning("Already connected to an instance");
            }

            var app = VA.Application.ApplicationHelper.FindRunningApplication();

            if (app == null)
            {
                throw new VA.AutomationException("Could not find an instance");
            }

            Globals.Application = app;

            this.WriteObject(app);
        }
    }
}