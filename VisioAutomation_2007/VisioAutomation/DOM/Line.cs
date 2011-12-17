using System.Collections.Generic;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.DOM
{
    public class Line : Shape
    {
        public VA.Drawing.Point P0 { get; private set; }
        public VA.Drawing.Point P1 { get; private set; }

        public Line(double x0, double y0, double x1, double y1)
        {
            this.P0 = new VA.Drawing.Point(x0, y0);
            this.P1 = new VA.Drawing.Point(x1, y1);
        }

        public Line(VA.Drawing.Point p0, VA.Drawing.Point p1)
        {
            this.P0 = p0;
            this.P1 = p1;
        }
    }
}