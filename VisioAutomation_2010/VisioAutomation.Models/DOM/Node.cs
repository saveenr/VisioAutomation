namespace VisioAutomation.Models.Dom;
public class Node
{
    public Node Parent { get; internal set; }
    public object Data { get; set; }

    protected Node()
    {
    }
}