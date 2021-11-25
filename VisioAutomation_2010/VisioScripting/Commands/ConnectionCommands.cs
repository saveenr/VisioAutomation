
using VisioAutomation.Shapes;


namespace VisioScripting.Commands;

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
    public List<VA.DocumentAnalysis.ConnectorEdge> GetTransitiveClosureOnActivePage(VisioScripting.TargetPage targetpage, VA.DocumentAnalysis.ConnectionAnalyzerOptions flag)
    {
        targetpage = targetpage.ResolveToPage(this._client);

        return VA.DocumentAnalysis.ConnectionAnalyzer.GetDirectedEdgesTransitive(targetpage.Page, flag);
    }

    public List<VA.DocumentAnalysis.ConnectorEdge> GetDirectedEdgesOnPage(TargetPage targetpage, VA.DocumentAnalysis.ConnectionAnalyzerOptions flag)
    {
        targetpage = targetpage.ResolveToPage(this._client);
        var directed_edges = VA.DocumentAnalysis.ConnectionAnalyzer.GetDirectedEdges(targetpage.Page, flag);
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