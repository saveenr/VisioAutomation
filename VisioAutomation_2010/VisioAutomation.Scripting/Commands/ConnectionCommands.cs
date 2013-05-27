using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class ConnectionCommands : CommandSet
    {
        private const string undoname_connectShapes = "Connect Shapes";

        public ConnectionCommands(Session session) :
            base(session)
        {

        }
        /// <summary>
        /// Returns all the connected pairs of shapes in the active page
        /// </summary>
        /// <param name="scripting_session"></param>
        /// <param name="flag"></param>
        /// <returns></returns>
        public IList<VA.Connections.ConnectorEdge> GetTransitiveClosure(Connections.ConnectorArrowEdgeHandling flag)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var app = this.Session.VisioApplication;
            return VA.Connections.PathAnalysis.GetTransitiveClosure(app.ActivePage, flag);
        }

        public IList<VA.Connections.ConnectorEdge> GetDirectedEdges(Connections.ConnectorArrowEdgeHandling flag)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var directed_edges = VA.Connections.PathAnalysis.GetEdges(this.Session.VisioApplication.ActivePage, flag);
            return directed_edges;
        }

        public IList<VA.Connections.ConnectorEdge> GetEdges()
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var edges = VA.Connections.PathAnalysis.GetEdges(this.Session.VisioApplication.ActivePage);
            this.Session.WriteVerbose( "{0} Edges found", edges.Count);
            return edges;
        }

        public IList<IVisio.Shape> Connect(IVisio.Master master, IList<IVisio.Shape> fromshapes, IList<IVisio.Shape> toshapes)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var active_page = this.Session.VisioApplication.ActivePage;

            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication, undoname_connectShapes))
            {
                var connectors = VA.Connections.ConnectorHelper.ConnectShapes(active_page, master, fromshapes, toshapes);
                return connectors;
            }
        }
    }
}