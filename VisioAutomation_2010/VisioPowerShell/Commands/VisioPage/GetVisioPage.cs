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

        [SMA.Parameter(Position=0, Mandatory = false)]
        public string Name;

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Document Document;


        [SMA.Parameter(Mandatory = false)] public VisioScripting.Models.PageType Type = PageType.Any;

        protected override void ProcessRecord()
        {
            if (this.ActivePage)
            {
                var page = this.Client.Page.GetActivePage();
                this.WriteObject(page);
                return;
            }

            // If the active page  is not specified then work on all the pages in a document (user-specified or auto)

            var targetdoc = new VisioScripting.TargetDocument(this.Document);
            var pages = this.Client.Page.FindPagesInDocumentByName(targetdoc, this.Name, this.Type);
            this.WriteObject(pages, true);
        }
    }
}