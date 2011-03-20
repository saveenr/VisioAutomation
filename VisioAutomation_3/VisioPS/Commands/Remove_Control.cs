using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Remove", "Control")]
    public class Remove_Control : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public int ControlIndex { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            this.ScriptingSession.Control.DeleteControl(this.ControlIndex);
        }
    }
}