using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.Connections
{
    public static class PathAnalysis
    {
        public static IList<ConnectorEdge> GetTransitiveClosure(
            IVisio.Page page,
            ConnectorEdgeHandling flag)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            var directed_edges = PathAnalysis.GetDirectedEdges(page, flag)
                .Select(e => new DirectedEdge<IVisio.Shape, IVisio.Shape>(e.From, e.To, e.Connector));

            var closure = PathAnalysis.GetClosureFromEdges(directed_edges)
                .Select(x => new ConnectorEdge(null, x.From, x.To)).ToList();

            return closure;
        }

        /// <summary>
        /// Returns all the directed,connected pairs of shapes in the  page
        /// </summary>
        /// <param name="page"></param>
        /// <param name="flag"></param>
        /// <returns></returns>
        public static IList<ConnectorEdge> GetDirectedEdges(
            IVisio.Page page,
            ConnectorEdgeHandling flag)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            var edges = PathAnalysis.GetDirectedEdgesRaw(page);

            if (flag == ConnectorEdgeHandling.Raw)
            {
                return edges;
            }

            // At this point we know we need to analyze the connetor arrows to produce the correct results

            var connnector_ids = edges.Select(e => e.Connector.ID).ToList();

            // Get the arrows for each connector
            var src_beginarrow = ShapeSheet.SRCConstants.BeginArrow;
            var src_endarrow = ShapeSheet.SRCConstants.EndArrow;

            var query = new ShapeSheet.Query.CellQuery();
            var col_beginarrow = query.AddCell(src_beginarrow, "BeginArrow");
            var col_endarrow = query.AddCell(src_endarrow, "EndArrow");

            var arrow_table = query.GetResults<int>(page, connnector_ids);
            
            IList<ConnectorEdge> directed_edges = new List<ConnectorEdge>();

            int connector_index = 0;
            foreach (var e in edges)
            {
                int beginarrow = arrow_table[connector_index][col_beginarrow];
                int endarrow = arrow_table[connector_index][col_endarrow];

                if ((beginarrow < 1) && (endarrow < 1))
                {
                    // the line has no arrows
                    if (flag == ConnectorEdgeHandling.Arrow_TreatConnectorsWithoutArrowsAsBidirectional)
                    {
                        // in this case treat the connector as pointing in both directions
                        var de1 = new ConnectorEdge(e.Connector, e.To, e.From);
                        var de2 = new ConnectorEdge(e.Connector, e.From, e.To);
                        directed_edges.Add(de1);
                        directed_edges.Add(de2);
                    }
                    else if (flag == ConnectorEdgeHandling.Arrow_ExcludeConnectorsWithoutArrows)
                    {
                        // in this case ignore the connector completely
                    }
                    else
                    {
                        throw new AutomationException("Internal error");
                    }
                }
                else
                {
                    // The connector has either a from-arrow, a to-arrow, or both

                    // handle if it has a from arrow
                    if (beginarrow > 0)
                    {
                        var de = new ConnectorEdge(e.Connector, e.To, e.From);
                        directed_edges.Add(de);
                    }

                    // handle if it has a to arrow
                    if (endarrow > 0)
                    {
                        var de = new ConnectorEdge(e.Connector, e.From, e.To);
                        directed_edges.Add(de);
                    }
                }

                connector_index++;
            }

            return directed_edges;
        }

        /// <summary>
        /// Gets all the pairs of shapes that are connected by a connector
        /// </summary>
        /// <param name="page"></param>
        /// <returns></returns>
        private static IList<ConnectorEdge> GetDirectedEdgesRaw(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            var page_connects = page.Connects;
            var connects = page_connects.AsEnumerable();

            var edges = new List<ConnectorEdge>();

            IVisio.Shape old_connect_shape = null;
            IVisio.Shape fromsheet = null;

            foreach (var connect in connects)
            {
                var current_connect_shape = connect.FromSheet;

                if (current_connect_shape != old_connect_shape)
                {
                    // the currect connector is NOT same as the one we stored previously
                    // this means the previous connector is connected to only one shape (not two).
                    // So skip the previos connector and start remembering from the current connector
                    old_connect_shape = current_connect_shape;
                    fromsheet = connect.ToSheet;
                }
                else
                {
                    // the currect connector is the same as the one we stored previously
                    // this means we have enountered it twice which means it connects two
                    // shapes and is thus an edge
                    var undirected_edge = new ConnectorEdge(current_connect_shape, fromsheet, connect.ToSheet);
                    edges.Add(undirected_edge);
                }
            }

            return edges;
        }

        internal static void PerformWarshall(Internal.BitArray2D adj_matrix)
        {
            if (adj_matrix == null)
            {
                throw new System.ArgumentNullException(nameof(adj_matrix));
            }

            if (adj_matrix.Width != adj_matrix.Height)
            {
                const string msg = "Adjacency Matrix width must equal its height";
                throw new System.ArgumentException(msg);
            }

            for (int k = 0; k < adj_matrix.Width; k++)
            {
                for (int row = 0; row < adj_matrix.Height; row++)
                {
                    for (int col = 0; col < adj_matrix.Width; col++)
                    {
                        bool v = adj_matrix.Get(row, col) | (adj_matrix.Get(row, k) & adj_matrix.Get(k, col));
                        adj_matrix[row, col] = v;
                    }
                }
            }
        }

        internal static IEnumerable<DirectedEdge<A, object>> GetClosureFromEdges<A, B>(
            IEnumerable<DirectedEdge<A, B>> edges)
        {
            if (edges == null)
            {
                throw new System.ArgumentNullException(nameof(edges));
            }

            var object_to_id = new Dictionary<A, int>();
            var id_to_object = new Dictionary<int, A>();

            foreach (var edge in edges)
            {
                if (!object_to_id.ContainsKey(edge.From))
                {
                    object_to_id[edge.From] = object_to_id.Count;
                }

                if (!object_to_id.ContainsKey(edge.To))
                {
                    object_to_id[edge.To] = object_to_id.Count;
                }
            }

            foreach (var i in object_to_id)
            {
                id_to_object[i.Value] = i.Key;
            }

            var internal_edges = new List<DirectedEdge<int, object>>();

            foreach (var edge in edges)
            {
                int fromid = object_to_id[edge.From];
                int toid = object_to_id[edge.To];
                var directed_edge = new DirectedEdge<int, object>(fromid, toid, null);
                internal_edges.Add(directed_edge);
            }

            if (internal_edges.Count == 0)
            {
                yield break;
            }

            int num_vertices = object_to_id.Count;
            var adj_matrix = new Internal.BitArray2D(num_vertices, num_vertices);
            foreach (var iedge in internal_edges)
            {
                adj_matrix[iedge.From, iedge.To] = true;
            }

            var warshall_result = adj_matrix.Clone();

            PathAnalysis.PerformWarshall(warshall_result);

            for (int row = 0; row < adj_matrix.Width; row++)
            {
                for (int col = 0; col < adj_matrix.Height; col++)
                {
                    if (warshall_result.Get(row, col) && (row!=col))
                    {
                        var de = new DirectedEdge<A, object>(id_to_object[row], id_to_object[col], null);
                        yield return de;
                    }
                }
            }
        }
    }
}