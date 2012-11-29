using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Get", "VisioUserDefinedCell")]
    public class Get_VisioUserDefinedCell : VisioPS.VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var dic = scriptingsession.UserDefinedCell.GetUserDefinedCells();
            this.WriteObject(dic);
        }
    }
}