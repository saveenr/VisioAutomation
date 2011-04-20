using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class TransitiveCLosure_Test
    {
        [TestMethod]
        public void TestTransitiveClosure()
        {
            string input =
                @"
                v3->v3
                v2->v3
                v0->v1
                v1->v2
                v4->v2  ";

            List<VA.Connections.DirectedEdge<int, object>> edges;
            var parse = parse_graph(input, out edges);

            int num_vertices = parse.Count;
            var adj_matrix = new VA.Internal.BitArray2D(num_vertices, num_vertices);
            foreach (var e in edges)
            {
                adj_matrix[e.From, e.To] = true;
            }

            var warshall_result = adj_matrix.Clone();
            VA.Connections.PathAnalysis.PerformWarshall(warshall_result);
        }

        private static IDictionary<string, int> parse_graph(string input, out List<VA.Connections.DirectedEdge<int, object>> edges)
        {
            char[] seps = { '\n' };
            string[] lines =
                input.Trim().Split(seps, System.StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).Where(
                    s => s.Length > 0).ToArray();

            edges = new List<VA.Connections.DirectedEdge<int, object>>();
            var dic_vname_to_vindex = new Dictionary<string, int>();
            var dic_vindex_to_vname = new Dictionary<int, string>();
            int n = 0;
            foreach (string line in lines)
            {
                System.Console.WriteLine(" {0} ", line);
                string[] xseps = { "->" };
                string[] tokens = line.Split(xseps, System.StringSplitOptions.RemoveEmptyEntries);
                string from = tokens[0];
                string to = tokens[1];
                if (!dic_vname_to_vindex.ContainsKey(from))
                {
                    dic_vname_to_vindex.Add(from, n);
                    dic_vindex_to_vname.Add(n, from);
                    n++;
                }
                if (!dic_vname_to_vindex.ContainsKey(to))
                {
                    dic_vname_to_vindex.Add(to, n);
                    dic_vindex_to_vname.Add(n, to);
                    n++;
                }

                var new_edge = new VA.Connections.DirectedEdge<int, object>(dic_vname_to_vindex[from], dic_vname_to_vindex[to], null);
                edges.Add(new_edge);
            }
            return dic_vname_to_vindex;
        }
    }
}