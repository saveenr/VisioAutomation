using VA = VisioAutomation;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.DirectedGraph
{   
    public class Drawing
    {
        public ShapeList Shapes;
        public ConnectorList Connectors;

        public Shape AddShape(string id, string label, string stencil_name, string master_name)
        {
            var s0 = new Shape(id);
            s0.Label = label;
            s0.StencilName = stencil_name;
            s0.MasterName = master_name;

            this.Shapes.Add(id, s0);
            return s0;
        }

        public Connector Connect(string id, Shape from, Shape to)
        {
            return Connect(id, from, to, id, VA.Shapes.Connections.ConnectorType.RightAngle);
        }

        public Connector Connect(
            string id, 
            Shape from, 
            Shape to, 
            string label,
             VA.Shapes.Connections.ConnectorType type)
        {
            var new_connector = new Connector(from, to);
            new_connector.Label = label;
            new_connector.ConnectorType = type;

            this.Connectors.Add(id, new_connector);
            return new_connector;
        }

        public Drawing()
        {
            this.Shapes = new ShapeList();
            this.Connectors = new ConnectorList();
        }

        public void Render(IVisio.Page page, VisioLayoutOptions options)
        {
            var vr = new VisioRenderer();
            vr.Render(page, this, options);
        }

        public void Render(IVisio.Page page, VA.Models.DirectedGraph.MSAGLLayoutOptions options)
        {
            // This is MSAGL-based render
            VA.Models.DirectedGraph.MSAGLRenderer.Render(page, this, options);
        }
    }
}