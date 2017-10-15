using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Redo, VisioPowerShell.Commands.Nouns.Visio)]
    public class RedoVisio : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.Client.Undo.RedoLastAction();
        }
    }
}