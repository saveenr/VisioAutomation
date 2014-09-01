using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting
{
    public struct DrawingSurface
    {
        public Microsoft.Office.Interop.Visio.Page Page;
        public Microsoft.Office.Interop.Visio.Master Master;
        public Microsoft.Office.Interop.Visio.Shape Shape;

        public DrawingSurface(IVisio.Page page)
        {
            this.Page = page;
            this.Master = null;
            this.Shape = null;
        }

        public DrawingSurface(IVisio.Master master)
        {
            this.Page = null;
            this.Master = master;
            this.Shape = null;
        }


        public DrawingSurface(IVisio.Shape shape)
        {
            this.Page = null;
            this.Master = null;
            this.Shape = shape;
        }


        public IVisio.Shape PolyLine(IList<VA.Drawing.Point> points)
        {
            var doubles_array = VA.Drawing.Point.ToDoubles(points).ToArray();

            if (this.Master != null)
            {
                var shape = this.Master.DrawPolyline(doubles_array, 0);
                return shape;
            }
            else if (this.Page != null)
            {
                var shape = this.Page.DrawPolyline(doubles_array, 0);
                return shape;
            }
            else if (this.Shape != null)
            {
                var shape = this.Shape.DrawPolyline(doubles_array, 0);
                return shape;
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }
        }
    }
}