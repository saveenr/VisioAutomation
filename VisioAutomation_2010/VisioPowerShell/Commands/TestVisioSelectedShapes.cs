using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsDiagnostic.Test, VisioPowerShell.Commands.Nouns.VisioSelectedShapes)]
    public class TestVisioSelectedShapes: VisioCmdlet
    {
        // checks to see if we have any selected shapes
        protected override void ProcessRecord()
        {
            this.WriteObject(this.Client.Selection.HasShapes());
        }
    }
}