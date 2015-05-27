using System.Collections.Generic;
using System.Management.Automation;
using VACONNECT = VisioAutomation.Shapes.Connections;

namespace VisioPowerShell.Commands.Get
{
    [Cmdlet(VerbsCommon.Get, "VisioDirectedEdge")]
    public class Get_VisioDirectedEdge : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public SwitchParameter GetShapeObjects { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter Raw { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter TreatUndirectedAsBidirectional { get; set; }

        protected override void ProcessRecord()
        {
            var flag = this.get_DirectedEdgeHandling();
            var edges = this.client.Connection.GetDirectedEdges(flag);

            if (this.GetShapeObjects)
            {
                this.WriteObject(edges, false);
                return;
            }

            this.write_edges_with_shapeids(edges);
                
        }

        private VACONNECT.ConnectorEdgeHandling get_DirectedEdgeHandling()
        {
            var flag = VACONNECT.ConnectorEdgeHandling.Arrow_ExcludeConnectorsWithoutArrows;

            if (this.Raw)
            {
                flag = VACONNECT.ConnectorEdgeHandling.Raw;
            }
            else
            {
                flag = this.TreatUndirectedAsBidirectional ? 
                    VACONNECT.ConnectorEdgeHandling.Arrow_TreatConnectorsWithoutArrowsAsBidirectional 
                    : VACONNECT.ConnectorEdgeHandling.Arrow_ExcludeConnectorsWithoutArrows;
            }
            return flag;
        }

        private void write_edges_with_shapeids(IList<VACONNECT.ConnectorEdge> edges)
        {
            foreach (var edge in edges)
            {
                var e = new Model.DirectedEdge(
                    edge.From.ID,
                    edge.To.ID,
                    edge.Connector.ID
                    );
                this.WriteObject(e);
            }
        }
    }
}