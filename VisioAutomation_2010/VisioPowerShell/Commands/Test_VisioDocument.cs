using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsDiagnostic.Test, VisioPowerShell.Nouns.VisioDocument)]
    public class Test_VisioDocument: VisioCmdlet
    {
        // checks to see if we hae an active drawing open
        protected override void ProcessRecord()
        {
            this.WriteObject(this.Client.Document.HasActiveDocument);
        }
    }
}