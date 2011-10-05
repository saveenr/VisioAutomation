using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using IVisio=Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout.DirectedGraph
{
    public class Drawing
    {
        private Dictionary<string, Shape> shapes;
        private Dictionary<string, Connector> connectors;

        public Shape AddShape(string id, string label, string stencil_name, string master_name)
        {
            var s0 = new Shape(id);
            s0.Label = label;
            s0.StencilName = stencil_name;
            s0.MasterName = master_name;

            this.shapes.Add(id, s0);
            return s0;
        }

        public IEnumerable<Shape> Shapes
        {
            get
            {
                foreach (var kv in this.shapes)
                {
                    yield return kv.Value;
                }
            }
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

            this.connectors.Add(id, new_connector);
            return new_connector;
        }

        public IEnumerable<Connector> Connectors
        {
            get
            {
                foreach (var kv in this.connectors)
                {
                    yield return kv.Value;
                }
            }
        }

        public Drawing()
        {
            this.shapes = new Dictionary<string, Shape>();
            this.connectors = new Dictionary<string, Connector>();
        }

        public void Render(IVisio.Page page, MSAGLLayoutOptions options)
        {
            var renderer = new VA.Layout.MSAGL.DirectedGraphLayout();
            renderer.LayoutOptions = options;
            renderer._render(this, page);
            page.ResizeToFitContents(renderer.LayoutOptions.ResizeBorderWidth);
        }

        private bool try_get_shape(string id, ref Shape shape)
        {
            if (this.shapes.ContainsKey(id))
            {
                shape = this.shapes[id];
                return true;
            }
            else
            {
                return false;
            }
        }

        public Shape GetShape(string id)
        {
            Shape shape = null;
            if (this.try_get_shape(id, ref shape))
            {
                return shape;
            }

            string msg = string.Format("Could not find shape with id '{0}'", id);
            throw new System.InvalidOperationException(msg);
        }

        public Shape FindShape(string id)
        {
            Shape shape = null;
            if (this.try_get_shape(id, ref shape))
            {
                return shape;
            }

            return null;
        }

        private bool try_get_connector(string id, ref Connector connector)
        {
            if (this.connectors.ContainsKey(id))
            {
                connector = this.connectors[id];
                return true;
            }
            else
            {
                return false;
            }
        }

        public Connector GetConnector(string id)
        {
            Connector c = null;
            if (this.try_get_connector(id, ref c))
            {
                return c;
            }

            string msg = string.Format("Could not find connector with id '{0}'", id);
            throw new System.InvalidOperationException(msg);
        }

        public Connector FindConnector(string id)
        {
            Connector c = null;

            if (this.try_get_connector(id, ref c))
            {
                return c;
            }

            return null;
        }

        public IEnumerable<string> ShapeIDs
        {
            get
            {
                foreach (var kv in this.shapes)
                {
                    yield return kv.Key;
                }
            }
        }
    }
}