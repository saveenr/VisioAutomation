
namespace VisioPowerShell.Commands.Visio
{
    [SMA.Cmdlet(SMA.VerbsCommon.Undo, Nouns.VisioApplication)]
    public class UndoVisio : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.Client.Undo.UndoLastAction();
        }
    }
}