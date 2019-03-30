using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioApplication
{
    [SMA.Cmdlet(SMA.VerbsDiagnostic.Test, Nouns.VisioApplication)]
    public class TestVisioApplication: VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            bool valid_app = this.Client.Application.ValidateActiveApplication();
            this.WriteObject(valid_app);
        }
    }
}