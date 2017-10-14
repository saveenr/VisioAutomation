using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, VisioPowerShell.Commands.Nouns.VisioPage)]
    public class RemoveVisioPage : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false, Position=0, ValueFromPipeline = true)]
        public IVisio.Page[] Pages;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Renumber;

        protected override void ProcessRecord()
        {
            if (this.Pages == null)
            {
                this.WriteVerbose("No Page objects ");
                this.WriteVerbose("Removing the Active Page");
                var page = this.Client.Application.GetApplication().ActivePage;
                this.Client.Page.DeletePages(new[] { page }, this.Renumber);
                return;
            }

            if (this.Pages != null)
            {
                this.WriteVerbose("Removing the Page Objects");
                this.Client.Page.DeletePages(this.Pages, this.Renumber);                
            }
        }
    }
}