using System.Collections.Generic;
using VisioPowerShell.Models;
using SMA = System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands.VisioPage
{
    [SMA.Cmdlet(SMA.VerbsDiagnostic.Measure, Nouns.VisioPage)]
    public class MeasureVisioPage : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter GetShapeObjects { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Raw { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TreatUndirectedAsBidirectional { get; set; }

        protected override void ProcessRecord()
        {
            var flag = this.get_DirectedEdgeHandling();
            var edges = this.Client.Connection.GetDirectedEdgesOnActivePage(flag);

            if (this.GetShapeObjects)
            {
                this.WriteObject(edges, true);
                return;
            }

            this.write_edges_with_shapeids(edges);
                
        }

        private VA.DocumentAnalysis.ConnectorHandling get_DirectedEdgeHandling()
        {
            var flag = new VA.DocumentAnalysis.ConnectorHandling();
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

        private void write_edges_with_shapeids(IList<VA.DocumentAnalysis.ConnectorEdge> edges)
        {
            foreach (var edge in edges)
            {
                var e = new DirectedEdge(
                    edge.From.ID,
                    edge.To.ID,
                    edge.Connector.ID
                    );
                this.WriteObject(e);
            }
        }
    }
}