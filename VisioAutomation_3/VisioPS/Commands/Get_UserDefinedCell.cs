using VAS = VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Get", "UserDefinedCell")]
    public class Get_UserDefinedCell : VisioPS.VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var dic = scriptingsession.UserDefinedCell.GetUserDefinedCells();
            this.WriteObject(dic);
        }
    }
}