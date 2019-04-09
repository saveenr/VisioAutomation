using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.Visio
{
    [SMA.Cmdlet(SMA.VerbsCommon.Undo, Nouns.Visio)]
    public class UndoVisio : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var activeapp = new VisioScripting.TargetActiveApplication();
            this.Client.Undo.UndoLastAction(activeapp);
        }
    }
}