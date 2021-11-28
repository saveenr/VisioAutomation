using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Analyzers
{
    public static class ConnectionAnalyzer
    {

        /// <summary>
        /// Returns all the directed,connected pairs of shapes in the  page
        /// </summary>
        /// <param name="page"></param>
        /// <param name="options"></param>
        /// <returns></returns>
        public static List<ConnectorEdge> GetDirectedEdges(
            IVisio.Page page,
            ConnectionAnalyzerOptions options)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            var edges = ConnectionAnalyzer._get_directed_edges_raw(page);

            if (options.DirectionSource == DirectionSource.UseConnectionOrder)
            {
                return edges;
            }

            // At this point we know we need to analyze the connetor arrows to produce the correct results

            var connnector_ids = edges.Select(e => e.Connector.ID).ToList();

            // Get the arrows for each connector
            var src_beginarrow = VisioAutomation.Core.SrcConstants.LineBeginArrow;
            var src_endarrow = VisioAutomation.Core.SrcConstants.LineEndArrow;

            var query = new VASS.Query.CellQuery();
            var col_beginarrow = query.Columns.Add(src_beginarrow);
            var col_endarrow = query.Columns.Add(src_endarrow);
            var listof_connectorinfo = query.GetResults<int>(page , connnector_ids);
            
            var directed_edges = new List<ConnectorEdge>();

            int connector_index = 0;
            foreach (var edge in edges)
            {
                var connector_info = listof_connectorinfo[connector_index];
                int beginarrow = connector_info[col_beginarrow];
                int endarrow = connector_info[col_endarrow];

                if ((beginarrow < 1) && (endarrow < 1))
                {
                    // the line has no arrows
                    if (options.NoArrowsHandling == NoArrowsHandling.TreatEdgeAsBidirectional)
                    {
                        // in this case treat the connector as pointing in both directions
                        var de1 = new ConnectorEdge(edge.Connector, edge.To, edge.From);
                        var de2 = new ConnectorEdge(edge.Connector, edge.From, edge.To);
                        directed_edges.Add(de1);
                        directed_edges.Add(de2);
                    }
                    else if (options.NoArrowsHandling == NoArrowsHandling.ExcludeEdge)
                    {
                        // in this case ignore the connector completely
                    }
                    else
                    {
                        throw new System.ArgumentOutOfRangeException(nameof(options));
                    }
                }
                else
                {
                    // The connector has either a from-arrow, a to-arrow, or both

                    // handle if it has a from arrow
                    if (beginarrow > 0)
                    {
                        var de = new ConnectorEdge(edge.Connector, edge.To, edge.From);
                        directed_edges.Add(de);
                    }

                    // handle if it has a to arrow
                    if (endarrow > 0)
                    {
                        var de = new ConnectorEdge(edge.Connector, edge.From, edge.To);
                        directed_edges.Add(de);
                    }
                }

                connector_index++;
            }

            return directed_edges;
        }

        public static List<ConnectorEdge> GetDirectedEdgesTransitive(
            IVisio.Page page,
            ConnectionAnalyzerOptions options)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            var directed_edges = ConnectionAnalyzer.GetDirectedEdges(page, options)
                .Select(e => new DirectedEdge<IVisio.Shape, IVisio.Shape>(e.From, e.To, e.Connector));

            var closure = ConnectionAnalyzer.GetClosureFromEdges(directed_edges)
                .Select(x => new ConnectorEdge(null, x.From, x.To)).ToList();

            return closure;
        }


        /// <summary>
        /// Gets all the pairs of shapes that are connected by a connector
        /// </summary>
        /// <param name="page"></param>
        /// <returns></returns>
        private static List<ConnectorEdge> _get_directed_edges_raw(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            var connects = page.Connects.ToList();

            var edges = new List<ConnectorEdge>();

            IVisio.Shape old_connect_shape = null;
            IVisio.Shape fromsheet = null;

            foreach (var connect in connects)
            {
                var current_connect_shape = connect.FromSheet;

                if (current_connect_shape != old_connect_shape)
                {
                    // the current connector is NOT same as the one we stored previously
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

        internal static void PerformWarshall(BitArray2D adj_matrix)
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

        public static IEnumerable<DirectedEdge<TNode, object>> GetClosureFromEdges<TNode, TData>(
            IEnumerable<DirectedEdge<TNode, TData>> edges)
        {
            if (edges == null)
            {
                throw new System.ArgumentNullException(nameof(edges));
            }

            var dicof_obj_to_id = new Dictionary<TNode, int>();
            var dicof_id_to_obj = new Dictionary<int, TNode>();

            foreach (var edge in edges)
            {
                if (!dicof_obj_to_id.ContainsKey(edge.From))
                {
                    dicof_obj_to_id[edge.From] = dicof_obj_to_id.Count;
                }

                if (!dicof_obj_to_id.ContainsKey(edge.To))
                {
                    dicof_obj_to_id[edge.To] = dicof_obj_to_id.Count;
                }
            }

            foreach (var kv in dicof_obj_to_id)
            {
                dicof_id_to_obj[kv.Value] = kv.Key;
            }

            var internal_edges = new List<DirectedEdge<int, object>>();

            foreach (var edge in edges)
            {
                int fromid = dicof_obj_to_id[edge.From];
                int toid = dicof_obj_to_id[edge.To];
                object data = null;
                var directed_edge = new DirectedEdge<int, object>(fromid, toid, data);
                internal_edges.Add(directed_edge);
            }

            if (internal_edges.Count == 0)
            {
                yield break;
            }

            int num_vertices = dicof_obj_to_id.Count;
            var adj_matrix = new BitArray2D(num_vertices, num_vertices);
            foreach (var internal_edge in internal_edges)
            {
                adj_matrix[internal_edge.From, internal_edge.To] = true;
            }

            var warshall_result = adj_matrix.Clone();

            ConnectionAnalyzer.PerformWarshall(warshall_result);

            for (int row = 0; row < adj_matrix.Width; row++)
            {
                for (int col = 0; col < adj_matrix.Height; col++)
                {
                    if (warshall_result.Get(row, col) && (row!=col))
                    {
                        var de = new DirectedEdge<TNode, object>(dicof_id_to_obj[row], dicof_id_to_obj[col], null);
                        yield return de;
                    }
                }
            }
        }
    }
}