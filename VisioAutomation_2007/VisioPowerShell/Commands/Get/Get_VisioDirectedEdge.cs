using System.Collections.Generic;
using SMA = System.Management.Automation;
using VA=VisioAutomation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioDirectedEdge")]
    public class Get_VisioDirectedEdge : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter GetShapeObjects { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Raw { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TreatUndirectedAsBidirectional { get; set; }

        protected override void ProcessRecord()
        {
            var flag = get_DirectedEdgeHandling();
            var edges = this.client.Connection.GetDirectedEdges(flag);

            if (this.GetShapeObjects)
            {
                this.WriteObject(edges, false);
                return;
            }

            write_edges_with_shapeids(edges);
                
        }

        private VA.Shapes.Connections.ConnectorEdgeHandling get_DirectedEdgeHandling()
        {
            var flag = VA.Shapes.Connections.ConnectorEdgeHandling.Arrow_ExcludeConnectorsWithoutArrows;

            if (this.Raw)
            {
                flag = VA.Shapes.Connections.ConnectorEdgeHandling.Raw;
            }
            else
            {
                if (this.TreatUndirectedAsBidirectional)
                {
                    flag = VA.Shapes.Connections.ConnectorEdgeHandling.Arrow_TreatConnectorsWithoutArrowsAsBidirectional;
                }
                else
                {
                    flag = VA.Shapes.Connections.ConnectorEdgeHandling.Arrow_ExcludeConnectorsWithoutArrows;
                }
            }
            return flag;
        }

        private void write_edges_with_shapeids(IList<VA.Shapes.Connections.ConnectorEdge> edges)
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