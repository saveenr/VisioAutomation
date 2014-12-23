using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, "VisioPage")]
    public class Remove_VisioPage : VisioCmdlet
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
                var page = this.client.VisioApplication.ActivePage;
                this.client.Page.Delete(new[] { page }, this.Renumber);
                return;
            }

            if (this.Pages != null)
            {
                this.WriteVerbose("Removing the Page Objects");
                this.client.Page.Delete(this.Pages, this.Renumber);                
            }
        }
    }
}