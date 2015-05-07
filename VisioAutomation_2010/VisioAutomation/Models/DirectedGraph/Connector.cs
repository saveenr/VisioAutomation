using VACONNECT=VisioAutomation.Shapes.Connections;
using VA=VisioAutomation;

namespace VisioAutomation.Models.DirectedGraph
{
    public class Connector : Node
    {
        public Shape From { get; set; }
        public Shape To { get; set; }

        public VACONNECT.ConnectorType ConnectorType { get; set; }

        public Connector(Shape from, Shape to)
        {
            this.ConnectorType = VACONNECT.ConnectorType.Curved;
            this.From = from;
            this.To = to;
        }
    }
}