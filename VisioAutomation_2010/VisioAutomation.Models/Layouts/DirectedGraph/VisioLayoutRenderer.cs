

namespace VisioAutomation.Models.Layouts.DirectedGraph;

public class VisioLayoutRenderer
{
    public DirectedGraphStyling Styling;
    public VisioLayoutOptions LayoutOptions;

    public VisioLayoutRenderer()
    {
        this.Styling = new DirectedGraphStyling();
        this.LayoutOptions = new VisioLayoutOptions();
    }

    public void Render(IVisio.Page page, DirectedGraphLayout dglayout)
    {
        // This is Visio-based render - it does NOT use MSAGL
        if (page == null)
        {
            throw new System.ArgumentNullException(nameof(page));
        }

        if (this.Styling== null)
        {
            throw new System.ArgumentNullException(nameof(this.Styling));
        }

        var page_node = new Dom.Page();
        double x = 0;
        double y = 1;
        foreach (var shape in dglayout.Nodes)
        {
            var shape_nodes = page_node.Shapes.Drop(shape.MasterName, shape.StencilName, x, y);
            shape.DomNode = shape_nodes;
            shape.DomNode.Text = new VisioAutomation.Models.Text.Element(shape.Label);
            x += 1.0;
        }

        foreach (var connector in dglayout.Edges)
        {
            var connector_node = page_node.Shapes.Connect(this.Styling.EdgeMasterName, this.Styling.EdgeStencilName, connector.From.DomNode, connector.To.DomNode);
            connector.DomNode = connector_node;
            connector.DomNode.Text = new VisioAutomation.Models.Text.Element(connector.Label);
        }

        page_node.ResizeToFit = true;
        page_node.ResizeToFitMargin = new VisioAutomation.Geometry.Size(0.5, 0.5);
        if (this.LayoutOptions.VisioLayoutStyle != null)
        {
            page_node.Layout = this.LayoutOptions.VisioLayoutStyle;
        }
        page_node.Render(page);
    }
}