using System.Management.Automation;
using VisioScripting.Models;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Get, VisioPowerShell.Commands.Nouns.VisioPage)]
    public class GetVisioPage : VisioCmdlet
    {
        [Parameter(Position=0, Mandatory = false)]
        public string Name=null;

        [Parameter(Mandatory = false)] public SwitchParameter ActivePage;

        [Parameter(Mandatory = false)] public VisioScripting.Models.PageType Type = PageType.Any;

        protected override void ProcessRecord()
        {
            var application = this.Client.Application.Get();

            if (this.ActivePage)
            {
                var page = this.Client.Page.Get();
                this.WriteObject(page);
                return;
            }

            var pages = this.Client.Page.GetPagesByName(this.Name, this.Type);
            this.WriteObject(pages, false);
        }
    }
}