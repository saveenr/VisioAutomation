using System.Management.Automation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.Get
{
    [Cmdlet(SMA.VerbsCommon.Get, "VisioPage")]
    public class Get_VisioPage : VisioCmdlet
    {
        [Parameter(Position=0, Mandatory = false)]
        public string Name=null;

        [Parameter(Mandatory = false)] public SMA.SwitchParameter ActivePage;

        protected override void ProcessRecord()
        {
            var application = this.client.Application.Get();

            if (this.ActivePage)
            {
                var page = this.client.Page.Get();
                this.WriteObject(page);
                return;
            }

            var pages = this.client.Page.GetPagesByName(this.Name);
            this.WriteObject(pages, true);
        }
    }
}