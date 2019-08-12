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

            var targetpages = new VisioScripting.TargetPages(this.Page).ResolveToPages(this.Client);

            if (targetpages.Pages.Count < 1)
            {
                return;
            }

            var list_pagedim = VisioScripting.Models.PageDimensions.Get_PageDimensions(targetpages.Pages);

            this.WriteObject(list_pagedim,true);

        }


        private void foo()
        {
            /*

            var targetpage = new VisioScripting.TargetPage(this.Page);

            var options = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            options.NoArrowsHandling = VA.DocumentAnalysis.NoArrowsHandling.ExcludeEdge;

            options.DirectionSource = VA.DocumentAnalysis.DirectionSource.UseConnectorArrows;

            options.NoArrowsHandling = this.TreatUndirectedAsBidirectional ?
                VA.DocumentAnalysis.NoArrowsHandling.TreatEdgeAsBidirectional
                : VA.DocumentAnalysis.NoArrowsHandling.ExcludeEdge;
            var edges = this.Client.Connection.GetDirectedEdgesOnPage(targetpage, options);
            this.WriteObject(edges, false);
                */
        }
    }
}

