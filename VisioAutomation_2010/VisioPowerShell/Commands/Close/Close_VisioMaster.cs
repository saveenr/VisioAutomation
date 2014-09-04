using VisioAutomation;

namespace VisioPowerShell.Commands
{
    [System.Management.Automation.Cmdlet(System.Management.Automation.VerbsCommon.Close, "VisioMaster")]
    public class Close_VisioMaster : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.ScriptingSession.Master.CloseMasterEditing();
        }
    }
}