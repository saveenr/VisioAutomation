using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioPage
{
    [SMA.Cmdlet(SMA.VerbsDiagnostic.Measure, Nouns.VisioPage)]
    public class MeasureVisioPage : VisioCmdlet
    {
        /*
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TreatUndirectedAsBidirectional { get; set; }
        */

        // CONTEXT:PAGES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page[] Page;

        protected override void ProcessRecord()
        {
            var targetpages = new VisioScripting.TargetPages(this.Page);
            var list_pagedim = this.Client.Page.GetPageDimensions(targetpages);
            this.WriteObject(list_pagedim,true);
        }
    }
}

