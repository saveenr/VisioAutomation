namespace VisioAutomation.Models.DirectedGraph
{
    public class NodeUserData
    {
        public string ID;
        public Shape Shape;
        public Connector Connector;

        public NodeUserData(string id, Shape shape)
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

        public NodeUserData(string id, Connector con)
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