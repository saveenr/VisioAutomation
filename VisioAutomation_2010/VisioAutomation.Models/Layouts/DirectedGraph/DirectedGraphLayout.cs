using VisioAutomation.Models.Dom;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class DirectedGraphLayout
    {
        public ShapeList Shapes;
        public ConnectorList Connectors;

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

        public Connector AddConnection(string id, Shape from, Shape to)
        {
            return this.AddConnection(id, from, to, id, ConnectorType.Default);
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

        public Connector AddConnection(string id, Shape from, Shape to, string label, string stencil_name, string master_name)
        {
            var new_connector = new Connector(from, to);
            new_connector.ID = id;
            new_connector.Label = label;
            new_connector.StencilName = stencil_name;
            new_connector.MasterName = master_name;

            this.Connectors.Add(id, new_connector);
            return new_connector;
        }

        public Connector AddConnection(string id, Shape from, Shape to, string label,
            ConnectorType type, int begin_arrow, int end_arrow, string hyperlink)
        {
            var new_connector = new Connector(from, to);
            new_connector.ID = id;
            new_connector.Label = label;
            new_connector.ConnectorType = type;
            new_connector.Cells = new ShapeCells();
            new_connector.Cells.LineBeginArrow = begin_arrow;
            new_connector.Cells.LineBeginArrowSize = begin_arrow;
            new_connector.Cells.LineEndArrow = end_arrow;
            new_connector.Cells.LineEndArrowSize = end_arrow;

            if (!string.IsNullOrEmpty(hyperlink))
            {


                //new_connector.VisioShape = IVisio.Shape; // IVisio.Shape();
                var h = new_connector.VisioShape.Hyperlinks.Add();

                h.Name = hyperlink; // Name of Hyperlink
                h.Address = hyperlink; // Address of Hyperlink
            }

            this.Connectors.Add(id, new_connector);
            return new_connector;
        }

        public void Render(IVisio.Page page, VisioLayoutOptions options)
        {
            var vr = new VisioRenderer();
            vr.Render(page, this, options);
        }

        public void Render(IVisio.Page page, MsaglLayoutOptions options)
        {
            MsaglRenderer.Render(page, this, options);
        }
    }
}