using System.Management.Automation;

namespace VisioPowerShell.Commands.Close
{
    [Cmdlet(VerbsCommon.Close, VisioPowerShell.Nouns.VisioMaster)]
    public class Close_VisioMaster : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.client.Master.CloseMasterEditing();
        }
    }
}