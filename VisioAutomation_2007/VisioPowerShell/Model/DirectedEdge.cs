namespace VisioPowerShell
{
    public class DirectedEdge
    {
        public readonly int FromShapeID;
        public readonly int ToShapeID;
        public readonly int ConnectorID;

        public DirectedEdge(int from, int to, int con)
        {
            this.FromShapeID = from;
            this.ToShapeID = to;
            this.ConnectorID = con;
        }
    }
}