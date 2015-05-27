using System.Management.Automation;

namespace VisioPowerShell.Commands.Test
{
    [Cmdlet(VerbsDiagnostic.Test, "VisioSelectedShapes")]
    public class Test_VisioSelectedShapes: VisioCmdlet
    {
        // checks to see if we have any selected shapes
        protected override void ProcessRecord()
        {
            this.WriteObject(this.client.Selection.HasShapes());
        }
    }
}