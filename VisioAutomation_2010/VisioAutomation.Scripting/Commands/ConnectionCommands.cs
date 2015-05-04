using System.Collections.Generic;
using VA_CXN=VisioAutomation.Shapes.Connections;
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
        public IList<VA_CXN.ConnectorEdge> GetTransitiveClosure(VA_CXN.ConnectorEdgeHandling flag)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var app = this.Client.VisioApplication;
            return VA_CXN.PathAnalysis.GetTransitiveClosure(app.ActivePage, flag);
        }

        public IList<VA_CXN.ConnectorEdge> GetDirectedEdges(VA_CXN.ConnectorEdgeHandling flag)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var directed_edges = VA_CXN.PathAnalysis.GetDirectedEdges(this.Client.VisioApplication.ActivePage, flag);
            return directed_edges;
        }

        public IList<IVisio.Shape> Connect(IList<IVisio.Shape> fromshapes, IList<IVisio.Shape> toshapes, IVisio.Master master)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var active_page = this.Client.VisioApplication.ActivePage;

            using (var undoscope = new Application.UndoScope(this.Client.VisioApplication, undoname_connectShapes))
            {
                if (master == null)
                {
                    var connectors = VA_CXN.ConnectorHelper.ConnectShapes(active_page, fromshapes, toshapes, null, false);
                    return connectors;                    
                }
                else
                {
                    var connectors = VA_CXN.ConnectorHelper.ConnectShapes(active_page, fromshapes, toshapes, master);
                    return connectors;
                }
            }
        }
    }
}