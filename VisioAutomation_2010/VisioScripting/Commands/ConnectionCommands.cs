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
        public List<VA.DocumentAnalysis.ConnectorEdge> GetTransitiveClosure(VA.DocumentAnalysis.ConnectorHandling flag)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            return VA.DocumentAnalysis.ConnectionAnalyzer.GetTransitiveClosure(cmdtarget.Page, flag);
        }

        public List<VA.DocumentAnalysis.ConnectorEdge> GetDirectedEdges(VA.DocumentAnalysis.ConnectorHandling flag)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            var directed_edges = VA.DocumentAnalysis.ConnectionAnalyzer.GetDirectedEdges(cmdtarget.Page, flag);
            return directed_edges;
        }

        public List<IVisio.Shape> Connect(IList<IVisio.Shape> fromshapes, IList<IVisio.Shape> toshapes, IVisio.Master master)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);
            
            using (var undoscope = this._client.Application.NewUndoScope("Connect Shapes"))
            {
                if (master == null)
                {
                    var connectors = ConnectorHelper.ConnectShapes(cmdtarget.Page, fromshapes, toshapes, null, false);
                    return connectors;                    
                }
                else
                {
                    var connectors = ConnectorHelper.ConnectShapes(cmdtarget.Page, fromshapes, toshapes, master);
                    return connectors;
                }
            }
        }
    }
}