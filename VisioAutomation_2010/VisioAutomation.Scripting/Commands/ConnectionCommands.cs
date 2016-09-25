using System.Collections.Generic;
using VisioAutomation.Shapes.Connectors;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Commands
{
    public class ConnectionCommands : CommandSet
    {
        internal ConnectionCommands(Client client) :
            base(client)
        {

        }
        /// <summary>
        /// Returns all the connected pairs of shapes in the active page
        /// </summary>
        /// <param name="flag"></param>
        /// <returns></returns>
        public List<VA.DocumentAnalysis.ConnectorEdge> GetTransitiveClosure(VA.DocumentAnalysis.ConnectorHandling flag)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var app = this._client.Application.Get();
            return VA.DocumentAnalysis.ConnectionAnalyzer.GetTransitiveClosure(app.ActivePage, flag);
        }

        public List<VA.DocumentAnalysis.ConnectorEdge> GetDirectedEdges(VA.DocumentAnalysis.ConnectorHandling flag)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            var directed_edges = VA.DocumentAnalysis.ConnectionAnalyzer.GetDirectedEdges(application.ActivePage, flag);
            return directed_edges;
        }

        public List<IVisio.Shape> Connect(IList<IVisio.Shape> fromshapes, IList<IVisio.Shape> toshapes, IVisio.Master master)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            var active_page = application.ActivePage;

            using (var undoscope = this._client.Application.NewUndoScope("Connect Shapes"))
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