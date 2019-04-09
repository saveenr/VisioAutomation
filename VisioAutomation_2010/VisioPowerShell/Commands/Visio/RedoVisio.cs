using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.Visio
{
    [SMA.Cmdlet(SMA.VerbsCommon.Redo, Nouns.Visio)]
    public class RedoVisio : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var activeapp = new VisioScripting.TargetActiveApplication();
            this.Client.Undo.RedoLastAction(activeapp);
        }
    }
}