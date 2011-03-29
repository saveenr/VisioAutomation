using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Get", "Layer")]
    public class Get_Layer : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var masters = scriptingsession.Master.GetMasters();
            this.WriteObject(masters);
        }
    }
}