using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsDiagnostic.Test, "VisioDocument")]
    public class Test_VisioDocument: VisioPS.VisioPSCmdlet
    {
        // checks to see if we hae an active drawing open
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            this.WriteObject(scriptingsession.HasActiveDrawing);
        }
    }
}