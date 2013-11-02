using VisioAutomation.Shapes.Connections;
using SMA = System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioDirectedEdge")]
    public class Get_VisioDirectedEdge : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public SMA.SwitchParameter UndirectedToDirected { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            var treat_as_connected = UndirectedToDirected
                ? ConnectorArrowEdgeHandling.TreatNoArrowEdgesAsBidirectional
                : ConnectorArrowEdgeHandling.ExcludeNoArrowEdges;

            var edges = scriptingsession.Connection.GetDirectedEdges(treat_as_connected);

            foreach (var edge in edges)
            {
                var e = new DirectedEdge();
                e.FromShapeID = edge.From.ID;
                e.ToShapeID = edge.To.ID;
                e.ConnectorID = edge.Connector.ID;

                this.WriteObject(e);                
            }
        }
    }

    public class DirectedEdge
    {
        public int FromShapeID;
        public int ToShapeID;
        public int ConnectorID;
    }
}
