namespace VisioAutomation.DOM
{
    public class Node
    {
        public Node Parent { get; internal set; }
        public object Data { get; set; }

        protected Node()
        {
        }
    }
}