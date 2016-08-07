using System.Collections.Generic;
using VACONNECT = VisioAutomation.Shapes.Connections;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Commands
{
    public class ConnectionCommands : CommandSet
    {
        private const string undoname_connectShapes = "Connect Shapes";

        internal ConnectionCommands(Client client) :
            base(client)
        {

        }
        /// <summary>
        /// Returns all the connected pairs of shapes in the active page
        /// </summary>
        /// <param name="flag"></param>
        /// <returns></returns>
        public IList<VA.DocumentAnalysis.ConnectorEdge> GetTransitiveClosure(VA.DocumentAnalysis.ConnectorEdgeHandling flag)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var app = this._client.Application.Get();
            return VA.DocumentAnalysis.ConnectionPathAnalyzer.GetTransitiveClosure(app.ActivePage, flag);
        }

        public IList<VA.DocumentAnalysis.ConnectorEdge> GetDirectedEdges(VA.DocumentAnalysis.ConnectorEdgeHandling flag)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            var directed_edges = VA.DocumentAnalysis.ConnectionPathAnalyzer.GetDirectedEdges(application.ActivePage, flag);
            return directed_edges;
        }

        public IList<IVisio.Shape> Connect(IList<IVisio.Shape> fromshapes, IList<IVisio.Shape> toshapes, IVisio.Master master)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            var active_page = application.ActivePage;

            using (var undoscope = this._client.Application.NewUndoScope(ConnectionCommands.undoname_connectShapes))
            {
                if (master == null)
                {
                    var connectors = VACONNECT.ConnectorHelper.ConnectShapes(active_page, fromshapes, toshapes, null, false);
                    return connectors;                    
                }
                else
                {
                    var connectors = VACONNECT.ConnectorHelper.ConnectShapes(active_page, fromshapes, toshapes, master);
                    return connectors;
                }
            }
        }
    }
}