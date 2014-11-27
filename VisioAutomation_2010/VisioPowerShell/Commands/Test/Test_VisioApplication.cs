using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsDiagnostic.Test, "VisioApplication")]
    public class Test_VisioApplication: VisioCmdlet
    {
        // checks to see if we hae an active drawing open
        protected override void ProcessRecord()
        {
            var app = this.client.VisioApplication;

            bool valid_app = this.client.Application.Validate();
            this.WriteObject(valid_app);
        }
    }
}