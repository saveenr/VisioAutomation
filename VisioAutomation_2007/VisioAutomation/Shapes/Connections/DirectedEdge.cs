namespace VisioAutomation.Shapes.Connections
{
    public struct DirectedEdge<TNode, TData>
    {
        public TNode From { get; private set; }
        public TNode To { get; private set; }
        public TData Data { get; private set; }

        public DirectedEdge(TNode from, TNode to, TData data)
            : this()
        {
            this.From = from;
            this.To = to;
            this.Data = data;
        }

    }
}