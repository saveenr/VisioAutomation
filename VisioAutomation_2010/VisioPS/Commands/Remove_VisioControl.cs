using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, "VisioControl")]
    public class Remove_VisioControl : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public int ControlIndex { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            this.ScriptingSession.Control.Delete(this.ControlIndex);
        }
    }
}