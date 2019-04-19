using System.Collections.Generic;
using VisioPowerShell.Models;
using VisioScripting.Models;
using SMA = System.Management.Automation;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioPage
{
    [SMA.Cmdlet(SMA.VerbsDiagnostic.Measure, Nouns.VisioPage)]
    public class MeasureVisioPage : VisioCmdlet
    {

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page Page;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TreatUndirectedAsBidirectional { get; set; }

        protected override void ProcessRecord()
        {
            var targetpage = new VisioScripting.TargetPage(this.Page);

            var options = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            options.NoArrowsHandling = VA.DocumentAnalysis.NoArrowsHandling.ExcludeEdge;

            options.DirectionSource = VA.DocumentAnalysis.DirectionSource.UseConnectorArrows;
            options.NoArrowsHandling = this.TreatUndirectedAsBidirectional ?
                VA.DocumentAnalysis.NoArrowsHandling.TreatEdgeAsBidirectional
                : VA.DocumentAnalysis.NoArrowsHandling.ExcludeEdge;

            var edges = this.Client.Connection.GetDirectedEdgesOnPage(targetpage,options);
            this.WriteObject(edges, false);
        }
    }
}