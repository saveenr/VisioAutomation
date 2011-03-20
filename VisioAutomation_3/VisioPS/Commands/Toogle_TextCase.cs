using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Toggle", "TextCase")]
    public class Toogle_TextCase : VisioPSCmdlet
    {

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.Text.ToogleCase();
        }
    }
}