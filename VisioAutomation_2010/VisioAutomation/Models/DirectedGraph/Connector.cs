using VACON=VisioAutomation.Shapes.Connections;
using VA=VisioAutomation;

namespace VisioAutomation.Models.DirectedGraph
{
    public class Connector : Node
    {
        public Shape From { get; set; }
        public Shape To { get; set; }

        public VACON.ConnectorType ConnectorType { get; set; }

        public Connector(Shape from, Shape to)
        {
            ConnectorType = VACON.ConnectorType.Curved;
            this.From = from;
            this.To = to;
        }
    }
}