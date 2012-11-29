using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Get", "VisioConnectionPoint")]
    public class Get_VisioConnectionPoint : VisioPS.VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var dic = scriptingsession.ConnectionPoint.Get();
            this.WriteObject(dic);
        }
    }
}