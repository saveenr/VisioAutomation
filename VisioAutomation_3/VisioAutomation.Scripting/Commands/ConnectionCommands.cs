using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VAS = VisioAutomation.Scripting;

namespace VisioAutomation.Scripting.Commands
{
    public class ConnectionCommands : SessionCommands
    {
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
            if (!HasActiveDrawing())
            {
                return new List<VA.Connections.ConnectorEdge>(0);
            }
            var app = Application;
            return VA.Connections.PathAnalysis.GetTransitiveClosure(app.ActivePage, flag);
        }

        /// <summary>
        /// Returns all the connected pairs of shapes in the active page
        /// </summary>
        /// <param name="scripting_session"></param>
        /// <param name="flag"></param>
        /// <returns></returns>
        public IList<VA.Connections.ConnectorEdge> GetDirectedEdges(Connections.ConnectorArrowEdgeHandling flag)
        {
            if (!HasActiveDrawing())
            {
                return new List<VA.Connections.ConnectorEdge>(0);
            }

            if (HasActiveDrawing())
            {
                var directed_edges = VA.Connections.PathAnalysis.GetEdges(Application.ActivePage, flag);
                return directed_edges;
            }
            else
            {
                return new List<VA.Connections.ConnectorEdge>(0);
            }
        }

        /// <summary>
        /// Returns all the connected pairs of shapes in the active page
        /// </summary>
        /// <param name="scripting_session"></param>
        /// <returns></returns>
        public IList<VA.Connections.ConnectorEdge> GetEdges()
        {
            IList<VA.Connections.ConnectorEdge> edges = new List<VA.Connections.ConnectorEdge>(0);

            if (HasActiveDrawing())
            {
                edges = VA.Connections.PathAnalysis.GetEdges(Application.ActivePage);
            }

            this.Session.Write(OutputStream.Verbose,"{0} Edges found", edges.Count);
            return edges;
        }

        public IList<IVisio.Shape> ConnectShapes(IVisio.Master master)
        {
            if (!HasSelectedShapes(2))
            {
                return new List<IVisio.Shape>(0);
            }

            var shapes = this.Session.Selection.EnumSelectedShapes().ToList();

            if (shapes.Count <= 1)
            {
                return new List<IVisio.Shape>(0);
            }

            var from_shapes = new List<IVisio.Shape>(shapes.Count);
            var to_shapes = new List<IVisio.Shape>(shapes.Count);
            var edges = SelectPairsOverlapped(shapes);

            foreach (var edge in edges)
            {
                from_shapes.Add(edge.From);
                to_shapes.Add(edge.To);
            }

            var active_page = this.Application.ActivePage;

            using (var undoscope = Application.CreateUndoScope())
            {
                var connectors = VA.Connections.ConnectorHelper.ConnectShapes(active_page, master, from_shapes, to_shapes);
                return connectors;
            }
        }

        /// <summary>
        /// Given an enumeration of returns them back as overlapping pairs
        /// </summary>
        /// <example>
        /// given input of (1,2,3,4,5,6,7,8)
        /// yields (1,2) (2,3), (3,4), (4,5), (5,6) (6,7), (7,8)
        /// </example>
        /// <param name="values">int input values</param>
        /// <returns>an enumeration of coordinates</returns>
        private static IEnumerable<VA.Connections.DirectedEdge<T, object>> SelectPairsOverlapped<T>(IEnumerable<T> values)
        {

            if (values == null)
            {
                throw new System.ArgumentNullException("values");
            }


            int count = 0;

            T first_value = default(T);
            foreach (var value in values)
            {
                if (count > 0)
                {
                    yield return new VA.Connections.DirectedEdge<T, object>(first_value, value, null);
                }
                first_value = value;
                count++;
            }
        }
        
        public IList<IVisio.Shape> ConnectShapes(IVisio.Master master, IList<IVisio.Shape> fromshapes, IList<IVisio.Shape> toshapes)
        {
            if (!HasActiveDrawing())
            {
                new List<IVisio.Shape>(0);
            }
            
            var application = Application;
            var active_page = this.Application.ActivePage;

            using (var undoscope = application.CreateUndoScope())
            {
                var connectors = VA.Connections.ConnectorHelper.ConnectShapes(active_page, master, fromshapes, toshapes);
                return connectors;
            }
        }

        public IList<IVisio.Shape> GetSelectedShapes(ShapesEnumeration enumerationtype)
        {
            if (!HasSelectedShapes())
            {
                return new List<IVisio.Shape>(0);
            }

            var selection = this.Session.Selection.GetSelection();
            return VA.SelectionHelper.GetSelectedShapes(selection, enumerationtype);
        }
    }
}