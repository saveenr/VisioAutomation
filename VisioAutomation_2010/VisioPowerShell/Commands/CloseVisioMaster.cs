using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Close, VisioPowerShell.Commands.Nouns.VisioMaster)]
    public class CloseVisioMaster : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.Client.Master.CloseMasterEditing();
        }
    }
}