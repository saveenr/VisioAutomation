using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using IVisio=Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Collections;

namespace VisioAutomation.Layout.DirectedGraph
{
    public class IDList<T> : IEnumerable<T> where T : class
    {
        private Dictionary<string, T> items;

        public IDList()
        {
            this.items  = new Dictionary<string, T>();
        }

        public void Add(string id, T g)
        {
            this.items.Add(id,g);
        }

        public T this[string index]
        {
            get { return this.items[index]; }
        }

        public int Count
        {
            get { return this.items.Count; }
        }

        public IEnumerator<T> GetEnumerator()
        {
            foreach (var i in this.items.Values)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()     // Explicit implementation
        {                                           // keeps it hidden.
            return GetEnumerator();
        }

        public bool ContainsKey(string id)
        {
            return this.items.ContainsKey(id);
        }

        public IEnumerable<string> IDs
        {
            get
            {
                foreach (var id in this.items.Keys)
                {
                    yield return id;
                }
                
            }
        }

        public T Find(string id)
        {
            T item = null;
            if (this.items.TryGetValue(id, out item))
            {
                return item;
            }

            return null;
        }        
    }

    public class ShapeList : IDList<Shape>
    {
        public ShapeList()
            : base()
        {
        }
    }

    public class ConnectorList : IDList<Connector>
    {
        public ConnectorList()
            : base()
        {
        }
    }


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

        public void Render(IVisio.Page page, MSAGLLayoutOptions options)
        {
            var renderer = new VA.Layout.MSAGL.DirectedGraphLayout();
            renderer.LayoutOptions = options;
            renderer._render(this, page);
            page.ResizeToFitContents(renderer.LayoutOptions.ResizeBorderWidth);
        }
    }
}