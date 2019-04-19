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
        public SMA.SwitchParameter Raw { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TreatUndirectedAsBidirectional { get; set; }

        protected override void ProcessRecord()
        {
            var targetpage = new VisioScripting.TargetPage(this.Page);

            var flag = this._get_directed_edge_handling();
            var edges = this.Client.Connection.GetDirectedEdgesOnPage(targetpage,flag);
            this.WriteObject(edges, false);
        }

        private VA.DocumentAnalysis.ConnectionAnalyzerOptions _get_directed_edge_handling()
        {
            var flag = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            flag.NoArrowsHandling =  VA.DocumentAnalysis.NoArrowsHandling.ExcludeEdge;

            if (this.Raw)
            {
                flag.DirectionSource = VA.DocumentAnalysis.DirectionSource.UseConnectionOrder;
            }
            else
            {
                flag.DirectionSource = VA.DocumentAnalysis.DirectionSource.UseConnectorArrows;
                flag.NoArrowsHandling = this.TreatUndirectedAsBidirectional ?
                    VA.DocumentAnalysis.NoArrowsHandling.TreatEdgeAsBidirectional 
                    : VA.DocumentAnalysis.NoArrowsHandling.ExcludeEdge;
            }
            return flag;
        }
    }
}