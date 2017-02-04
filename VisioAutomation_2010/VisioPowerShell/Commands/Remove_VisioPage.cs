using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Remove, VisioPowerShell.Nouns.VisioPage)]
    public class Remove_VisioPage : VisioCmdlet
    {
        [Parameter(Mandatory = false, Position=0, ValueFromPipeline = true)]
        public IVisio.Page[] Pages;

        [Parameter(Mandatory = false)]
        public SwitchParameter Renumber;

        protected override void ProcessRecord()
        {
            if (this.Pages == null)
            {
                this.WriteVerbose("No Page objects ");
                this.WriteVerbose("Removing the Active Page");
                var page = this.Client.Application.Get().ActivePage;
                this.Client.Page.Delete(new[] { page }, this.Renumber);
                return;
            }

            if (this.Pages != null)
            {
                this.WriteVerbose("Removing the Page Objects");
                this.Client.Page.Delete(this.Pages, this.Renumber);                
            }
        }
    }
}