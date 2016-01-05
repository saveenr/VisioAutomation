using System.Management.Automation;

namespace VisioPowerShell.Commands.Get
{
    [Cmdlet(VerbsCommon.Get, VisioPowerShell.Nouns.VisioPage)]
    public class Get_VisioPage : VisioCmdlet
    {
        [Parameter(Position=0, Mandatory = false)]
        public string Name=null;

        [Parameter(Mandatory = false)] public SwitchParameter ActivePage;

        protected override void ProcessRecord()
        {
            var application = this.Client.Application.Get();

            if (this.ActivePage)
            {
                var page = this.Client.Page.Get();
                this.WriteObject(page);
                return;
            }

            var pages = this.Client.Page.GetPagesByName(this.Name);
            this.WriteObject(pages, true);
        }
    }
}