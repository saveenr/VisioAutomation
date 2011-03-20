using VAS = VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Get", "Control")]
    public class Get_Control : VisioPSCmdlet
    {

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var controls = this.ScriptingSession.Control.GetControls();

            this.WriteObject(controls);
        }
    }
}