using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Shapes.Connections;
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
            return Connect(id, from, to, id, ConnectorType.RightAngle);
        }

        public Connector Connect(
            string id, 
            Shape from, 
            Shape to, 
            string label,
             ConnectorType type)
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

            var page_node = new VA.DOM.Page();
            double x = 0;
            double y = 1;
            foreach (var shape in this.Shapes)
            {
                var shape_nodes = page_node.Shapes.Drop(shape.MasterName, shape.StencilName, x, y);
                shape.DOMNode = shape_nodes;
                shape.DOMNode.Text = new VA.Text.Markup.TextElement( shape.Label ) ;
                x += 1.0;
            }

            foreach (var connector in this.Connectors)
            {

                var connector_node = page_node.Shapes.Connect("Dynamic Connector", "basic_u.vss", connector.From.DOMNode, connector.To.DOMNode);
                connector.DOMNode = connector_node;
                connector.DOMNode.Text = new VA.Text.Markup.TextElement( connector.Label );
            }
            page_node.ResizeToFit = true;
            page_node.ResizeToFitMargin = new VA.Drawing.Size(0.5, 0.5);
            page_node.Render(page);
        }

        public void Render(IVisio.Page page, VA.Models.DirectedGraph.MSAGLLayoutOptions options)
        {
            VA.Models.DirectedGraph.MSAGLRenderer.Render(page, this, options);
        }
    }
}