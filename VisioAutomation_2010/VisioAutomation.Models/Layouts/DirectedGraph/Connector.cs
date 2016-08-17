using VisioAutomation.Shapes.Connectors;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class Connector : Node
    {
        public Shape From { get; set; }
        public Shape To { get; set; }

        public ConnectorType ConnectorType { get; set; }
	 
        public System.Collections.Generic.List<Dom.Hyperlink> Hyperlinks { get; set; }
        public string StencilName { get; set; }
        public string MasterName { get; set; }

        public Connector(Shape from, Shape to)
        {
            this.ConnectorType = ConnectorType.Curved;
            this.From = from;
            this.To = to;
        }
    }
}