using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsDiagnostic.Test, Nouns.VisioShape)]
    public class TestVisioShape: VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var something_is_selected = this.Client.Selection.ContainsShapes(new VisioScripting.TargetSelection());
            this.WriteObject(something_is_selected);
        }
    }
}