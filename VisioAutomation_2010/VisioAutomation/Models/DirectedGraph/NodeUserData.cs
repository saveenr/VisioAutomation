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
                throw new System.ArgumentNullException(nameof(id));
            }

            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            this.ID = id;
            this.Shape = shape;
        }

        public NodeUserData(string id, Connector con)
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
            this.Connector = con;
        }
    }
}