using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsDiagnostic.Test, "VisioDocument")]
    public class Test_VisioDocument: VisioCmdlet
    {
        // checks to see if we hae an active drawing open
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            this.WriteObject(scriptingsession.HasActiveDocument);
        }
    }
}