using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class DirectedGraphLayout
    {
        public readonly IDList<Node> Nodes;
        public readonly IDList<Edge> Edges;

        public DirectedGraphLayout()
        {
            this.Nodes = new IDList<Node>();
            this.Edges = new IDList<Edge>();
        }

        public Node AddNode(string id, string label, string stencil_name, string master_name)
        {
            var new_node = new Node(id);
            new_node.Label = label;
            new_node.StencilName = stencil_name;
            new_node.MasterName = master_name;

            this.Nodes.Add(id, new_node);
            return new_node;
        }

        public Edge AddEdge(
            string id,
            Node from,
            Node to,
            string label,
            ConnectorType type)
        {
            var new_edge = new Edge(from, to);
            new_edge.ID = id;
            new_edge.Label = label;
            new_edge.ConnectorType = type;

            this.Edges.Add(id, new_edge);
            return new_edge;
        }
    }
}