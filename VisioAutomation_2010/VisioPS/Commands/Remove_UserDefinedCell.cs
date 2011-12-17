using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Remove", "UserDefinedCell")]
    public class Remove_UserDefinedCell : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string Name { get; set; }
        
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.UserDefinedCell.Delete(Name);
        }
    }
}