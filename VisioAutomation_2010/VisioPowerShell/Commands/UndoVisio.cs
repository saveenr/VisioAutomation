using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Undo, VisioPowerShell.Commands.Nouns.Visio)]
    public class UndoVisio : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.Client.Application.Undo();
        }
    }
}