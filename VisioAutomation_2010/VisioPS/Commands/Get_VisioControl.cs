using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Get", "VisioControl")]
    public class Get_VisioControl : VisioPSCmdlet
    {

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var controls = this.ScriptingSession.Control.Get();

            this.WriteObject(controls);
        }
    }
}