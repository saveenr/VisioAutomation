namespace VisioPowerShell.Commands.VisioShape;

[SMA.Cmdlet(SMA.VerbsDiagnostic.Test, Nouns.VisioShape)]
public class TestVisioShape: VisioCmdlet
{
    protected override void ProcessRecord()
    {
        var something_is_selected = this.Client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto);
        this.WriteObject(something_is_selected);
    }
}