using VA=VisioAutomation;

namespace VisioAutomation.Layout.Models.DirectedGraph
{
    public class Connector : Node
    {
        public Shape From { get; set; }
        public Shape To { get; set; }

        public VA.Connections.ConnectorType ConnectorType { get; set; }

        public Connector(Shape from, Shape to)
        {
            ConnectorType = VA.Connections.ConnectorType.Curved;
            this.From = from;
            this.To = to;
        }
    }
}