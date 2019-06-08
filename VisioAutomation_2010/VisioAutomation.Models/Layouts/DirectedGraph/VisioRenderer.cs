using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    class VisioRenderer
    {
        public VisioRenderer()
        {
        }

        public void Render(IVisio.Page page, DirectedGraphLayout dglayout, DirectedGraphStyling dgstyling, VisioLayoutOptions visiooptions)
        {
            // This is Visio-based render - it does NOT use MSAGL
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            if (dgstyling== null)
            {
                throw new System.ArgumentNullException(nameof(dgstyling));
            }

            var page_node = new Dom.Page();
            double x = 0;
            double y = 1;
            foreach (var shape in dglayout.Shapes)
            {
                var shape_nodes = page_node.Shapes.Drop(shape.MasterName, shape.StencilName, x, y);
                shape.DomNode = shape_nodes;
                shape.DomNode.Text = new VisioAutomation.Models.Text.Element(shape.Label);
                x += 1.0;
            }

            foreach (var connector in dglayout.Connectors)
            {
                var connector_node = page_node.Shapes.Connect(dgstyling.EdgeMasterName, dgstyling.EdgeStencilName, connector.From.DomNode, connector.To.DomNode);
                connector.DomNode = connector_node;
                connector.DomNode.Text = new VisioAutomation.Models.Text.Element(connector.Label);
            }

            page_node.ResizeToFit = true;
            page_node.ResizeToFitMargin = new VisioAutomation.Geometry.Size(0.5, 0.5);
            if (visiooptions.VisioLayoutStyle != null)
            {
                page_node.Layout = visiooptions.VisioLayoutStyle;
            }
            page_node.Render(page);
        }
    }
}