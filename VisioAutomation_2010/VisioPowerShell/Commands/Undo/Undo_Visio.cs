using System.Management.Automation;

namespace VisioPowerShell.Commands.Undo
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