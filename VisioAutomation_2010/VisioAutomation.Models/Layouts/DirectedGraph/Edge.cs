namespace VisioAutomation.Models.Layouts.DirectedGraph;

public class Edge : Element
{
    public Node From { get; set; }
    public Node To { get; set; }

    public string StencilName { get; set; }
    public string MasterName { get; set; }
    public ConnectorType ConnectorType { get; set; }

    public Edge(Node from, Node to)
    {
        this.From = from;
        this.To = to;
        this.ConnectorType = ConnectorType.Curved;
    }
}