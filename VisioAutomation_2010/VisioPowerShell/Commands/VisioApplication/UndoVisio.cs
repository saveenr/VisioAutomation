using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Undo, VisioPowerShell.Commands.Nouns.Visio)]
    public class UndoVisio : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.Client.Undo.UndoLastAction();
        }
    }
}