namespace VisioAutomation.DocumentAnalysis;

public struct DirectedEdge<TNode, TData>
{
    public readonly TNode From;
    public readonly TNode To;
    public readonly TData Data;

    public DirectedEdge(TNode from, TNode to, TData data)
        : this()
    {
        this.From = from;
        this.To = to;
        this.Data = data;
    }
}