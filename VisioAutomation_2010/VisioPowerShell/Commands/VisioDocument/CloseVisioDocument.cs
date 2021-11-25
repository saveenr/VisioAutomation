
namespace VisioPowerShell.Commands.VisioDocument
{
    [SMA.Cmdlet(SMA.VerbsCommon.Close, Nouns.VisioDocument)]
    public class CloseVisioDocument : VisioCmdlet
    {
        // CONTEXT:DOCUMENTS
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Document[] Document;

        protected override void ProcessRecord()
        {
            var targetdocs = new VisioScripting.TargetDocuments(this.Document);
            this.Client.Document.CloseDocument(targetdocs);
        }
    }
}