using System;
using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Layout.Models.DirectedGraph
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
            return Connect(id, from, to, id, VA.Connections.ConnectorType.RightAngle);
        }

        public Connector Connect(
            string id, 
            Shape from, 
            Shape to, 
            string label,
             VA.Connections.ConnectorType type)
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

        public void Render(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            var dompage = new VA.DOM.Page();
            double x = 0;
            double y = 1;
            foreach (var shape in this.Shapes)
            {
                var dom_node = dompage.Shapes.Drop(shape.MasterName, shape.StencilName, x, y);
                shape.DOMNode = dom_node;
                shape.DOMNode.Text = new VA.Text.Markup.TextElement( shape.Label ) ;
                x += 1.0;
            }

            foreach (var connector in this.Connectors)
            {

                var dom_node = dompage.Shapes.Connect("Dynamic Connector", "basic_u.vss", connector.From.DOMNode, connector.To.DOMNode);
                connector.DOMNode = dom_node;
                connector.DOMNode.Text = new VA.Text.Markup.TextElement( connector.Label );
            }
            dompage.ResizeToFit = true;
            dompage.ResizeToFitMargin = new VA.Drawing.Size(0.5, 0.5);
            dompage.Render(page);
        }

        public void Render(IVisio.Page page, VA.Layout.Models.DirectedGraph.MSAGLLayoutOptions options)
        {
            VA.Layout.Models.DirectedGraph.MSAGLRenderer.Render(page, this, options);
        }
    }
}