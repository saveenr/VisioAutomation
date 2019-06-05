using VisioAutomation.Models.Dom;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class DirectedGraphLayout
    {
        public readonly ShapeList Shapes;
        public readonly ConnectorList Connectors;

        public DirectedGraphLayout()
        {
            this.Shapes = new ShapeList();
            this.Connectors = new ConnectorList();
        }

        public Shape AddShape(string id, string label, string stencil_name, string master_name)
        {
            var s0 = new Shape(id);
            s0.Label = label;
            s0.StencilName = stencil_name;
            s0.MasterName = master_name;

            this.Shapes.Add(id, s0);
            return s0;
        }

        public Connector AddConnection(
            string id,
            Shape from,
            Shape to,
            string label,
            ConnectorType type)
        {
            var new_connector = new Connector(from, to);
            new_connector.ID = id;
            new_connector.Label = label;
            new_connector.ConnectorType = type;

            this.Connectors.Add(id, new_connector);
            return new_connector;
        }

        public void Render(IVisio.Page page, DirectedGraphStyling dgstyling)
        {
            var vr = new VisioRenderer();
            vr.Render(page, this, dgstyling);
        }

        public void Render(IVisio.Page page, MsaglLayoutOptions layoutoptions)
        {
            MsaglRenderer.Render(page, this, layoutoptions);
        }
    }
}