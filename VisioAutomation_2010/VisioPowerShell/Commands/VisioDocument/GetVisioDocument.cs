using VisioPowerShell.Models;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioDocument
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioDocument)]
    public class GetVisioDocument : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ActiveDocument;

        [SMA.Parameter(Position = 0, Mandatory = false)]
        public string Name = null;
        
        
        protected override void ProcessRecord()
        {
            if (this.ActiveDocument)
            {
                var application = this.Client.Application.GetAttachedApplication();
                var active_doc = application.ActiveDocument;
                this.WriteObject(active_doc);
                return;
            }

            // If the active document is not specified then work on all the pages in the application

            var docs = this.Client.Document.FindDocuments(this.Name);
            this.WriteObject(docs, true);
        }
    }


}