using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommunications.Connect, "VisioApplication")]
    public class Connect_VisioApplication : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var app = scriptingsession.Application.Attach();

            if (app == null)
            {
                throw new VA.AutomationException("Could not find an instance of the Visio Application");
            }

            this.WriteVerboseEx("Attaching to Visio Application");
            this.WriteObject(app);
        }
    }
}