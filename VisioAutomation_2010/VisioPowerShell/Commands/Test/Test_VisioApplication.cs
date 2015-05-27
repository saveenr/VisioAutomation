using System.Management.Automation;

namespace VisioPowerShell.Commands.Test
{
    [Cmdlet(VerbsDiagnostic.Test, "VisioApplication")]
    public class Test_VisioApplication: VisioCmdlet
    {
        // checks to see if we hae an active drawing open
        protected override void ProcessRecord()
        {
            bool valid_app = this.client.Application.Validate();
            this.WriteObject(valid_app);
        }
    }
}