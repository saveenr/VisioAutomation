 using System.Collections.Generic;
using VisioAutomation.Shapes;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
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
        public List<VA.Analyzers.ConnectorEdge> GetTransitiveClosureOnActivePage(VisioScripting.TargetPage targetpage, VA.Analyzers.ConnectionAnalyzerOptions flag)
        {
            targetpage = targetpage.ResolveToPage(this._client);

            return VA.Analyzers.ConnectionAnalyzer.GetDirectedEdgesTransitive(targetpage.Page, flag);
        }

        public List<VA.Analyzers.ConnectorEdge> GetDirectedEdgesOnPage(TargetPage targetpage, VA.Analyzers.ConnectionAnalyzerOptions flag)
        {
            targetpage = targetpage.ResolveToPage(this._client);
            var directed_edges = VA.Analyzers.ConnectionAnalyzer.GetDirectedEdges(targetpage.Page, flag);
            return directed_edges;
        }

        public List<IVisio.Shape> ConnectShapes(VisioScripting.TargetPage target_page, IList<IVisio.Shape> fromshapes, IList<IVisio.Shape> toshapes, IVisio.Master master)
        {
            target_page = target_page.ResolveToPage(this._client);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(ConnectShapes)))
            {
                if (master == null)
                {
                    var connectors = ConnectorHelper.ConnectShapes(target_page.Page, fromshapes, toshapes, null, false);
                    return connectors;                    
                }
                else
                {
                    var connectors = ConnectorHelper.ConnectShapes(target_page.Page, fromshapes, toshapes, master);
                    return connectors;
                }
            }
        }
    }
}