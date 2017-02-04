using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Undo, VisioPowerShell.Nouns.Visio)]
    public class Undo_Visio : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.Client.Application.Undo();
        }
    }
}