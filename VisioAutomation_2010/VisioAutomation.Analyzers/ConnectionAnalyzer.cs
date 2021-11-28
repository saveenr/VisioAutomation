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

            if (options.EdgeDirectionSource == EdgeDirectionSource.UseConnectionOrder)
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
                    if (options.EdgeNoArrowsHandling == EdgeNoArrowsHandling.IncludeEdgesForBothDirections)
                    {
                        // in this case treat the connector as pointing in both directions
                        var de1 = new ConnectorEdge(edge.Connector, edge.To, edge.From);
                        var de2 = new ConnectorEdge(edge.Connector, edge.From, edge.To);
                        directed_edges.Add(de1);
                        directed_edges.Add(de2);
                    }
                    else if (options.EdgeNoArrowsHandling == EdgeNoArrowsHandling.ExcludeEdge)
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

            // First cache how objects and ids relate
            // because we will be looking this up a lot later

            var obj_to_id = new Dictionary<TNode, int>();
            var id_to_obj = new Dictionary<int, TNode>();

            foreach (var edge in edges)
            {
                if (!obj_to_id.ContainsKey(edge.From))
                {
                    obj_to_id[edge.From] = obj_to_id.Count;
                }

                if (!obj_to_id.ContainsKey(edge.To))
                {
                    obj_to_id[edge.To] = obj_to_id.Count;
                }
            }

            foreach (var kv in obj_to_id)
            {
                id_to_obj[kv.Value] = kv.Key;
            }

            // Create a collection to store all the edges we will discover

            var internal_edges = new List<DirectedEdge<int, object>>();

            // Add the initial input edges to the collection

            foreach (var edge in edges)
            {
                int fromid = obj_to_id[edge.From];
                int toid = obj_to_id[edge.To];
                object data = null;
                var directed_edge = new DirectedEdge<int, object>(fromid, toid, data);
                internal_edges.Add(directed_edge);
            }

            // If there are are no edges at at this point, there is nothing left to do

            if (internal_edges.Count == 0)
            {
                yield break;
            }

            // Construct the initial adjacency matrix

            int num_vertices = obj_to_id.Count;
            var adj_matrix = new BitArray2D(num_vertices, num_vertices);
            foreach (var internal_edge in internal_edges)
            {
                adj_matrix[internal_edge.From, internal_edge.To] = true;
            }

            // Clone the adjacency matrix, and fill it in 
            // with the transitive closure as specified from the Warshall algortihm

            var warshall_result = adj_matrix.Clone();
            ConnectionAnalyzer.PerformWarshall(warshall_result);

            // For each item in the where an transitive closure is indicated
            // create a directed edge object
            for (int row = 0; row < adj_matrix.Width; row++)
            {
                for (int col = 0; col < adj_matrix.Height; col++)
                {
                    if (warshall_result.Get(row, col) && (row!=col))
                    {
                        var de = new DirectedEdge<TNode, object>(id_to_obj[row], id_to_obj[col], null);
                        yield return de;
                    }
                }
            }
        }
    }
}