using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommunications.Connect, "VisioApplication")]
    public class Connect_VisioApplication : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            if (scriptingsession.VisioApplication!= null)
            {
                this.WriteWarning("Already connected to an instance");
            }

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