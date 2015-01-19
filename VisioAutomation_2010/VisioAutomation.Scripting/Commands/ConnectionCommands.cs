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

        public ConnectionCommands(Client client) :
            base(client)
        {

        }
        /// <summary>
        /// Returns all the connected pairs of shapes in the active page
        /// </summary>
        /// <param name="flag"></param>
        /// <returns></returns>
        public IList<ConnectorEdge> GetTransitiveClosure(ConnectorEdgeHandling flag)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var app = this.Client.VisioApplication;
            return PathAnalysis.GetTransitiveClosure(app.ActivePage, flag);
        }

        public IList<ConnectorEdge> GetDirectedEdges(ConnectorEdgeHandling flag)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var directed_edges = PathAnalysis.GetDirectedEdges(this.Client.VisioApplication.ActivePage, flag);
            return directed_edges;
        }

        public IList<IVisio.Shape> Connect(IList<IVisio.Shape> fromshapes, IList<IVisio.Shape> toshapes, IVisio.Master master)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var active_page = this.Client.VisioApplication.ActivePage;

            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication, undoname_connectShapes))
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