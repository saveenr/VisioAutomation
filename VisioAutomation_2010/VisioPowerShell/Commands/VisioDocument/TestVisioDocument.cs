using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioDocument
{
    [SMA.Cmdlet(SMA.VerbsDiagnostic.Test, Nouns.VisioDocument)]
    public class TestVisioDocument: VisioCmdlet
    {
        // checks to see if we hae an active drawing open
        protected override void ProcessRecord()
        {
            this.WriteObject(this.Client.Document.HasActiveDocument);
        }
    }
}