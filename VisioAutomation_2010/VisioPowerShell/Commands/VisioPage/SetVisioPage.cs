using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioPage
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, Nouns.VisioPage)]
    public class SetVisioPage : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "Page")]
        public IVisio.Page Page  { get; set; }

        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "Flags")]
        public VisioScripting.Models.PageRelativePosition RelativePosition { get; set; }
        
        protected override void ProcessRecord()
        {
            if (this.Page != null)
            {
                var targetpage = new VisioScripting.TargetPage(this.Page);
                this.Client.Page.SetActivePage(targetpage);
            }
            else
            {
                this.Client.Page.SetActivePage(VisioScripting.TargetDocument.Active, this.RelativePosition);                
            }
        }
    }
}