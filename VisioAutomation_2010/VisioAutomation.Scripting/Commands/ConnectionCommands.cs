using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VisioAutomation.Shapes.Connections;
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
        public IList<ConnectorEdge> GetTransitiveClosure(ConnectorArrowEdgeHandling flag)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var app = this.Session.VisioApplication;
            return PathAnalysis.GetTransitiveClosure(app.ActivePage, flag);
        }

        public IList<ConnectorEdge> GetDirectedEdges(ConnectorArrowEdgeHandling flag)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var directed_edges = PathAnalysis.GetEdges(this.Session.VisioApplication.ActivePage, flag);
            return directed_edges;
        }

        public IList<ConnectorEdge> GetEdges()
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var edges = PathAnalysis.GetEdges(this.Session.VisioApplication.ActivePage);
            this.Session.WriteVerbose( "{0} Edges found", edges.Count);
            return edges;
        }

        public IList<IVisio.Shape> Connect(IList<IVisio.Shape> fromshapes, IList<IVisio.Shape> toshapes, IVisio.Master master)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var active_page = this.Session.VisioApplication.ActivePage;

            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication, undoname_connectShapes))
            {
                if (master == null)
                {
                    var connectors = ConnectorHelper.ConnectShapes(active_page, fromshapes, toshapes, null, false);
                    return connectors;                    
                }
                else
                {
                    var connectors = ConnectorHelper.ConnectShapes(active_page, fromshapes, toshapes, master);
                    return connectors;
                }
            }
        }
    }
}