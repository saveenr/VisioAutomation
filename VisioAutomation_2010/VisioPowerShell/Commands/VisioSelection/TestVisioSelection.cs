using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsDiagnostic.Test, Nouns.VisioSelection)]
    public class TestVisioSelection: VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var something_is_selected = this.Client.Selection.SelectionContainsShapes();
            this.WriteObject(something_is_selected);
        }
    }
}