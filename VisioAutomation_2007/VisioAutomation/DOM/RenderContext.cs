using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.DOM
{
    internal class RenderContext
    {
        private Dictionary<short, IVisio.Shape> id_to_shape;
        private IVisio.Shapes pageshapes;
        public IVisio.Page VisioPage { get; private set; }

        public RenderContext(IVisio.Page visio_page)
        {
            this.id_to_shape = new Dictionary<short, IVisio.Shape>();
            this.VisioPage = visio_page;
            this.pageshapes = visio_page.Shapes;
        }

        public IVisio.Shape GetShape(short id)
        {
            IVisio.Shape vshape;
            if (this.id_to_shape.TryGetValue(id, out vshape))
            {
                return vshape;
            }
            else
            {
                vshape = this.pageshapes.ItemFromID16[id];
                this.id_to_shape[id] = vshape;
                return vshape;
            }
        }
    }
}