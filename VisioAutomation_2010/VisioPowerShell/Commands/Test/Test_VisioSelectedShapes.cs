using System.Management.Automation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.Test
{
    [Cmdlet(SMA.VerbsDiagnostic.Test, "VisioSelectedShapes")]
    public class Test_VisioSelectedShapes: VisioCmdlet
    {
        // checks to see if we have any selected shapes
        protected override void ProcessRecord()
        {
            this.WriteObject(this.client.Selection.HasShapes());
        }
    }
}