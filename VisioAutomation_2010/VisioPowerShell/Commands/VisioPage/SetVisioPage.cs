using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioPage
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, Nouns.VisioPage)]
    public class SetVisioPage : VisioCmdlet
    {
        // NONCONTEXT:SHAPE

        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "Page")]
        public IVisio.Page Page  { get; set; }

        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "Flags")]
        public VisioScripting.Models.PageRelativePosition RelativePosition { get; set; }
        
        protected override void ProcessRecord()
        {
            if (this.Page != null)
            {
                this.Client.Page.SetActivePage(this.Page);
            }
            else
            {
                this.Client.Page.SetActivePage(VisioScripting.TargetDocument.Auto, this.RelativePosition);                
            }
        }
    }
}