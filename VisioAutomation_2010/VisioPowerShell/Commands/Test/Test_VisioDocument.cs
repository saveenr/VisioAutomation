using System.Management.Automation;

namespace VisioPowerShell.Commands.Test
{
    [Cmdlet(VerbsDiagnostic.Test, "VisioDocument")]
    public class Test_VisioDocument: VisioCmdlet
    {
        // checks to see if we hae an active drawing open
        protected override void ProcessRecord()
        {
            this.WriteObject(this.client.Document.HasActiveDocument);
        }
    }
}