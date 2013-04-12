using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class ConnectionCommands : CommandSet
    {
        string undoname_connectShapes = "Connect Shapes";

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
            this.CheckApplication();
            if (!this.Session.HasActiveDrawing)
            {
                return new List<VA.Connections.ConnectorEdge>(0);
            }
            var app = this.Session.VisioApplication;
            return VA.Connections.PathAnalysis.GetTransitiveClosure(app.ActivePage, flag);
        }

        public IList<VA.Connections.ConnectorEdge> GetDirectedEdges(Connections.ConnectorArrowEdgeHandling flag)
        {
            this.CheckApplication();
            if (!this.Session.HasActiveDrawing)
            {
                return new List<VA.Connections.ConnectorEdge>(0);
            }

            if (this.Session.HasActiveDrawing)
            {
                var directed_edges = VA.Connections.PathAnalysis.GetEdges(this.Session.VisioApplication.ActivePage, flag);
                return directed_edges;
            }
            else
            {
                return new List<VA.Connections.ConnectorEdge>(0);
            }
        }

        public IList<VA.Connections.ConnectorEdge> GetEdges()
        {
            this.CheckApplication();
            IList<VA.Connections.ConnectorEdge> edges = new List<VA.Connections.ConnectorEdge>(0);

            if (this.Session.HasActiveDrawing)
            {
                edges = VA.Connections.PathAnalysis.GetEdges(this.Session.VisioApplication.ActivePage);
            }

            this.Session.WriteVerbose( "{0} Edges found", edges.Count);
            return edges;
        }

        public IList<IVisio.Shape> Connect(IVisio.Master master, IList<IVisio.Shape> fromshapes, IList<IVisio.Shape> toshapes)
        {
            this.CheckApplication();
            if (!this.Session.HasActiveDrawing)
            {
                new List<IVisio.Shape>(0);
            }

            var active_page = this.Session.VisioApplication.ActivePage;

            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication, undoname_connectShapes))
            {
                var connectors = VA.Connections.ConnectorHelper.ConnectShapes(active_page, master, fromshapes, toshapes);
                return connectors;
            }
        }
    }
}