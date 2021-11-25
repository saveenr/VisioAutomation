namespace VisioPowerShell.Commands.Visio;

[SMA.Cmdlet(SMA.VerbsCommon.Redo, Nouns.VisioApplication)]
public class RedoVisio : VisioCmdlet
{
    protected override void ProcessRecord()
    {
        this.Client.Undo.RedoLastAction();
    }
}