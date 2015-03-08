namespace VisioAutomation.Models.DirectedGraph
{
    public class NodeUserData
    {
        public string ID;
        public VisioAutomation.Models.DirectedGraph.Shape Shape;
        public VisioAutomation.Models.DirectedGraph.Connector Connector;

        public NodeUserData(string id, VisioAutomation.Models.DirectedGraph.Shape shape)
        {
            if (id == null)
            {
                throw new System.ArgumentNullException("id");
            }

            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            this.ID = id;
            this.Shape = shape;
        }

        public NodeUserData(string id, VisioAutomation.Models.DirectedGraph.Connector con)
        {
            if (id == null)
            {
                throw new System.ArgumentNullException("id");
            }

            if (con == null)
            {
                throw new System.ArgumentNullException("con");
            }

            this.ID = id;
            this.Connector = con;
        }
    }
}