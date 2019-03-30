using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioMaster
{
    [SMA.Cmdlet(SMA.VerbsCommon.Close, Nouns.VisioMaster)]
    public class CloseVisioMaster : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.Client.Master.CloseMasterEditing();
        }
    }
}