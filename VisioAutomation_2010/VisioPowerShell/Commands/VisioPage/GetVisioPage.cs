using VisioScripting.Models;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioPage
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioPage)]
    public class GetVisioPage : VisioCmdlet
    {
        [SMA.Parameter(Position=0, Mandatory = false)]
        public string Name=null;

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Document Document;

        [SMA.Parameter(Mandatory = false)] public 
        SMA.SwitchParameter ActivePage;

        [SMA.Parameter(Mandatory = false)] public VisioScripting.Models.PageType Type = PageType.Any;

        protected override void ProcessRecord()
        {
            if (this.ActivePage)
            {
                var page = this.Client.Page.GetActivePage();
                this.WriteObject(page);
                return;
            }

            var targetdoc = new TargetDocument(this.Document).Resolve(this.Client);
            var pages = this.Client.Page.FindPagesInDocumentByName(targetdoc, this.Name, this.Type);
            this.WriteObject(pages, true);
        }
    }
}