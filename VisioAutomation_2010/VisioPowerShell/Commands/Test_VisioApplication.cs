using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsDiagnostic.Test, VisioPowerShell.Commands.Nouns.VisioApplication)]
    public class Test_VisioApplication: VisioCmdlet
    {
        // checks to see if we hae an active drawing open
        protected override void ProcessRecord()
        {
            bool valid_app = this.Client.Application.Validate();
            this.WriteObject(valid_app);
        }
    }
}