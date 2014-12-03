using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommunications.Connect, "VisioApplication")]
    public class Connect_VisioApplication : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var app = this.client.Application.Attach();

            if (app == null)
            {
                throw new VA.AutomationException("Could not find an instance of the Visio Application");
            }

            this.WriteVerbose("Attaching to Visio Application");
            this.WriteObject(app);
        }
    }
}