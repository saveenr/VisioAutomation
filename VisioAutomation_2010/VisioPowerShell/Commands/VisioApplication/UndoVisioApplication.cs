using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioApplication
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