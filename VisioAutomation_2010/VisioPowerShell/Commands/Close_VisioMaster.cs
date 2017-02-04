using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Close, VisioPowerShell.Nouns.VisioMaster)]
    public class Close_VisioMaster : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.Client.Master.CloseMasterEditing();
        }
    }
}