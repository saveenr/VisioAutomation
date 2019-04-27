using VisioScripting.Models;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioPage
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioPage)]
    public class GetVisioPage : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ActivePage;

        [SMA.Parameter(Mandatory = false)]
        public int[] ID;

        [SMA.Parameter(Position=0, Mandatory = false)]
        public string Name;

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Document Document;
        
        protected override void ProcessRecord()
        {
            if (this.ActivePage)
            {
                var page_active = this.Client.Page.GetActivePage();
                this.WriteObject(page_active);
                return;
            }

            // If the active page  is not specified then work on all the pages in a document (user-specified or auto)

            var targetdoc = new VisioScripting.TargetDocument(this.Document);

            // First, the ID case
            if (this.ID != null)
            {
                var t = targetdoc.ResolveToDocument(this.Client);
                foreach (var id in this.ID)
                {
                    var page = t.Document.Pages[id];
                    this.WriteObject(page);
                }
                return;
            }

            // Then, handle the name case
            if (this.Name!=null)
            {
                var pages_by_name = this.Client.Page.FindPagesInDocument(targetdoc, this.Name);
                this.WriteObject(pages_by_name, true);
                return;
            }

            // Finally return all the pages in the document
            var pages_in_doc = this.Client.Page.FindPagesInDocument(targetdoc, null);
            this.WriteObject(pages_in_doc, true);
            return;

        }
    }
}