namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class ElementUserData
    {
        public readonly string ID;
        public readonly Node Node;
        public readonly Edge Edge;

        public ElementUserData(string id, Node node)
        {
            this.ID = id ?? throw new System.ArgumentNullException(nameof(id));
            this.Node = node ?? throw new System.ArgumentNullException(nameof(node));
        }

        public ElementUserData(string id, Edge con)
        {
            this.ID = id ?? throw new System.ArgumentNullException(nameof(id));
            this.Edge = con ?? throw new System.ArgumentNullException(nameof(con));
        }
    }
}