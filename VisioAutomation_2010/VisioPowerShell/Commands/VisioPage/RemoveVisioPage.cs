using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, Nouns.VisioPage)]
    public class RemoveVisioPage : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false, Position=0, ValueFromPipeline = true)]
        public IVisio.Page[] Pages;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Renumber;

        protected override void ProcessRecord()
        {
            var target_pages = new VisioScripting.Models.TargetPages(this.Pages);
            target_pages.Resolve(this.Client);
            this.Client.Page.DeletePages(target_pages, this.Renumber);
        }
    }
}