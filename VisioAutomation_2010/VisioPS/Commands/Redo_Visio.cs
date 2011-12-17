using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Redo", "Visio")]
    public class Redo_Visio : VisioPS.VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            scriptingsession.Application.Redo();
        }
    }
}