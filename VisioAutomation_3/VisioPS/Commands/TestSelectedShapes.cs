using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Test", "SelectedShapes")]
    public class TestSelectedShapes: VisioPS.VisioPSCmdlet
    {
        // checks to see if we have any selected shapes
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            this.WriteObject(scriptingsession.HasSelectedShapes());
        }
    }
}