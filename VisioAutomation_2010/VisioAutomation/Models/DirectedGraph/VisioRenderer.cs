namespace VisioAutomation.Models.DirectedGraph
{
    class VisioRenderer
    {
        public static void Render(Microsoft.Office.Interop.Visio.Page page, VisioAutomation.Models.DirectedGraph.Drawing drawing, VisioLayoutOptions options)
        {
            // This is Visio-based render - it does NOT use MSAGL
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            var page_node = new VisioAutomation.DOM.Page();
            double x = 0;
            double y = 1;
            foreach (var shape in drawing.Shapes)
            {
                var shape_nodes = page_node.Shapes.Drop(shape.MasterName, shape.StencilName, x, y);
                shape.DOMNode = shape_nodes;
                shape.DOMNode.Text = new VisioAutomation.Text.Markup.TextElement(shape.Label);
                x += 1.0;
            }

            foreach (var connector in drawing.Connectors)
            {

                var connector_node = page_node.Shapes.Connect("Dynamic Connector", "basic_u.vss", connector.From.DOMNode, connector.To.DOMNode);
                connector.DOMNode = connector_node;
                connector.DOMNode.Text = new VisioAutomation.Text.Markup.TextElement(connector.Label);
            }
            page_node.ResizeToFit = true;
            page_node.ResizeToFitMargin = new VisioAutomation.Drawing.Size(0.5, 0.5);
            page_node.Render(page);
        }
    }
}