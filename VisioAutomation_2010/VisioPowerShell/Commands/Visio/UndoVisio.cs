using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.Visio
{
    [SMA.Cmdlet(SMA.VerbsCommon.Undo, Nouns.Visio)]
    public class UndoVisio : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.Client.Undo.UndoLastAction(new VisioScripting.TargetActiveApplication());
        }
    }
}