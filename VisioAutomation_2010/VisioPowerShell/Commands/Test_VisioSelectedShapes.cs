using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsDiagnostic.Test, VisioPowerShell.Nouns.VisioSelectedShapes)]
    public class Test_VisioSelectedShapes: VisioCmdlet
    {
        // checks to see if we have any selected shapes
        protected override void ProcessRecord()
        {
            this.WriteObject(this.Client.Selection.HasShapes());
        }
    }
}