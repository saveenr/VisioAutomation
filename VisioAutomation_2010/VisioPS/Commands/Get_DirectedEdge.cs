using SMA = System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Get", "DirectedEdge")]
    public class Get_DirectedEdge : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 1, Mandatory = false)]
        public VisioAutomation.Connections.ConnectorArrowEdgeHandling TreatAsConnected { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var edges = scriptingsession.Connection.GetDirectedEdges(TreatAsConnected);
            this.WriteObject(edges);
        }
    }
}
