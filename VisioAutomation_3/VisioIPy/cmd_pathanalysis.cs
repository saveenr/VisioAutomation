using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public IList<VA.Connections.ConnectorEdge> GetTransitiveClosure(bool treatasconnected)
        {
            var flag = treatasconnected
                           ? VA.Connections.PathAnalysis.ConnectorArrowEdgeHandling.TreatNoArrowEdgesAsBidirectional
                           : VA.Connections.PathAnalysis.ConnectorArrowEdgeHandling.ExcludeNoArrowEdges;
            return this.ScriptingSession.Connection.GetTransitiveClosure(flag);
        }

        public IList<VA.Connections.ConnectorEdge> GetEdgesFromConnections()
        {
            return this.ScriptingSession.Connection.GetEdges();
        }

        public IList<VA.Connections.ConnectorEdge> GetDirectedEdgesFromConnections(bool treat_edges_as_bidirectional)
        {
            var flag = treat_edges_as_bidirectional
                           ? VA.Connections.PathAnalysis.ConnectorArrowEdgeHandling.TreatNoArrowEdgesAsBidirectional
                           : VA.Connections.PathAnalysis.ConnectorArrowEdgeHandling.ExcludeNoArrowEdges;
            var pairs = this.ScriptingSession.Connection.GetDirectedEdges(flag);
            return pairs;
        }
    }
}