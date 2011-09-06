using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.DOM
{
    public class PieSlice: Shape
    {
        public VA.Drawing.Point Center { get; private set; }
        public double Radius { get; private set; }
        public double Start { get; private set; }
        public double End  { get; private set; }

        public PieSlice(double x0, double y0, double r, double start, double end)
        {
            this.Center = new VA.Drawing.Point(x0, y0);
            this.Radius= r;
            this.Start = start;
            this.End = end;
        }

        public PieSlice(VA.Drawing.Point p0, double r, double start, double end)
        {
            this.Center = p0;
            this.Radius = r;
            this.Start = start;
            this.End = end;
        }
    }
}