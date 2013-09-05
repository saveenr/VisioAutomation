using VisioAutomation.Shapes.Connections;
using SMA = System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioDirectedEdge")]
    public class Get_VisioDirectedEdge : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public ConnectorArrowEdgeHandling TreatAsConnected { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var edges = scriptingsession.Connection.GetDirectedEdges(TreatAsConnected);
            this.WriteObject(edges);
        }
    }
}
