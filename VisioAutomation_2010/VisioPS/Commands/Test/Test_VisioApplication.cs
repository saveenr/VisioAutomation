using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsDiagnostic.Test, "VisioApplication")]
    public class Test_VisioApplication: VisioCmdlet
    {
        // checks to see if we hae an active drawing open
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var app = scriptingsession.VisioApplication;

            bool valid_app = scriptingsession.Application.Validate();
            this.WriteObject(valid_app);
        }
    }
}