using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioPage
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, Nouns.VisioPage)]
    public class RemoveVisioPage : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Renumber;

        [SMA.Parameter(Mandatory = false, Position = 0, ValueFromPipeline = true)]
        public IVisio.Page[] Page;

        protected override void ProcessRecord()
        {
            var targetpages = new VisioScripting.TargetPages(this.Page);
            this.Client.Page.DeletePages(targetpages, this.Renumber);
        }
    }
}