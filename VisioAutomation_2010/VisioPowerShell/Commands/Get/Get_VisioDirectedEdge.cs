using System.Collections.Generic;
using VisioAutomation.Shapes.Connections;
using SMA = System.Management.Automation;
using VA=VisioAutomation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Get, "VisioDirectedEdge")]
    public class Get_VisioDirectedEdge : VisioCmdlet
    {
        [SMA.ParameterAttribute(Mandatory = false)]
        public SMA.SwitchParameter GetShapeObjects { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
        public SMA.SwitchParameter Raw { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
        public SMA.SwitchParameter TreatUndirectedAsBidirectional { get; set; }

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

        private ConnectorEdgeHandling get_DirectedEdgeHandling()
        {
            var flag = ConnectorEdgeHandling.Arrow_ExcludeConnectorsWithoutArrows;

            if (this.Raw)
            {
                flag = ConnectorEdgeHandling.Raw;
            }
            else
            {
                if (this.TreatUndirectedAsBidirectional)
                {
                    flag = ConnectorEdgeHandling.Arrow_TreatConnectorsWithoutArrowsAsBidirectional;
                }
                else
                {
                    flag = ConnectorEdgeHandling.Arrow_ExcludeConnectorsWithoutArrows;
                }
            }
            return flag;
        }

        private void write_edges_with_shapeids(IList<ConnectorEdge> edges)
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