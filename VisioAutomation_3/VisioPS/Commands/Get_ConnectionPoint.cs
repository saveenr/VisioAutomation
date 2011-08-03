using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Get", "ConnectionPoint")]
    public class Get_ConnectionPoint : VisioPS.VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var dic = scriptingsession.ConnectionPoint.Get();
            this.WriteObject(dic);
        }
    }
}