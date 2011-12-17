using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Get", "CustomProperty")]
    public class Get_CustomProperty : VisioPS.VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var dic = scriptingsession.CustomProp.GetCustomProperties();
            this.WriteObject(dic);
        }
    }
}