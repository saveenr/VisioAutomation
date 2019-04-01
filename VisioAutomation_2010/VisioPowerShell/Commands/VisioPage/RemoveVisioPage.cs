using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioPage
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
            var targetpages = new VisioScripting.Models.TargetPages(this.Pages);
            targetpages.Resolve(this.Client);
            this.Client.Page.DeletePages(targetpages, this.Renumber);
        }
    }
}