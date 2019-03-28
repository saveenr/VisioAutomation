using SMA = System.Management.Automation;
using VisioScripting.Models;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioPage)]
    public class GetVisioPage : VisioCmdlet
    {
        [SMA.Parameter(Position=0, Mandatory = false)]
        public string Name=null;

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

            var pages = this.Client.Page.FindPagesInActiveDocumentByName(this.Name, this.Type);
            this.WriteObject(pages, true);
        }
    }
}