using VACONNECT = VisioAutomation.Shapes.Connections;

namespace VisioAutomation.Models.DirectedGraph
{
    public class Connector : Node
    {
        public Shape From { get; set; }
        public Shape To { get; set; }

        public VACONNECT.ConnectorType ConnectorType { get; set; }
	 
        public System.Collections.Generic.List<DOM.Hyperlink> Hyperlinks { get; set; }
        public string StencilName { get; set; }
        public string MasterName { get; set; }

        public Connector(Shape from, Shape to)
        {
            this.ConnectorType = VACONNECT.ConnectorType.Curved;
            this.From = from;
            this.To = to;
        }
    }
}