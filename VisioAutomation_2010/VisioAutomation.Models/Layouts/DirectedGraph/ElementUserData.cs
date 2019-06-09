namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class ElementUserData
    {
        public string ID;
        public Node Node;
        public Edge Edge;

        public ElementUserData(string id, Node node)
        {
            if (id == null)
            {
                throw new System.ArgumentNullException(nameof(id));
            }

            if (node == null)
            {
                throw new System.ArgumentNullException(nameof(node));
            }

            this.ID = id;
            this.Node = node;
        }

        public ElementUserData(string id, Edge con)
        {
            if (id == null)
            {
                throw new System.ArgumentNullException(nameof(id));
            }

            if (con == null)
            {
                throw new System.ArgumentNullException(nameof(con));
            }

            this.ID = id;
            this.Edge = con;
        }
    }
}