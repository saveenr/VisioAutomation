using VisioScripting;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioDocument
{
    [SMA.Cmdlet(SMA.VerbsData.Save, Nouns.VisioDocument)]
    public class SaveVisioDocument : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        [SMA.ValidateNotNullOrEmpty]
        public string Filename;

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Document Document;

        protected override void ProcessRecord()
        {
            var targetdoc = new TargetDocument(this.Document);

            if (this.Filename!=null)
            {
                this.Client.Document.SaveDocumentAs(targetdoc, this.Filename);
            }
            else
            {
                this.Client.Document.SaveDocument(targetdoc);
            }
        }
    }
}