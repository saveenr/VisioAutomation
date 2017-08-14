using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsDiagnostic.Test, VisioPowerShell.Commands.Nouns.VisioSelectedShapes)]
    public class TestVisioSelectedShapes: VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var something_is_selected = this.Client.Selection.HasShapes();
            this.WriteObject(something_is_selected);
        }
    }
}