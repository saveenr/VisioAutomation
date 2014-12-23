using VisioAutomation.Shapes.Connections;
using VA=VisioAutomation;

namespace VisioAutomation.Models.DirectedGraph
{
    public class Connector : Node
    {
        public Shape From { get; set; }
        public Shape To { get; set; }

        public ConnectorType ConnectorType { get; set; }

        public Connector(Shape from, Shape to)
        {
            ConnectorType = ConnectorType.Curved;
            this.From = from;
            this.To = to;
        }
    }
}