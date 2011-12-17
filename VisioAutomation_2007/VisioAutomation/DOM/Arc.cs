using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.DOM
{
    public class Arc: Shape
    {
        public VA.Drawing.Point Center { get; private set; }
        public double InnerRadius { get; private set; }
        public double OuterRadius { get; private set; }
        public double StartAngle { get; private set; }
        public double EndAngle  { get; private set; }

        public Arc(double x0, double y0, double ri, double ro, double start, double end)
        {
            this.Center = new VA.Drawing.Point(x0, y0);
            this.InnerRadius= ri;
            this.OuterRadius= ro;
            this.StartAngle = start;
            this.EndAngle = end;
        }

        public Arc(VA.Drawing.Point p0, double ri, double ro, double start, double end)
        {
            this.Center = p0;
            this.InnerRadius = ri;
            this.OuterRadius = ro;
            this.StartAngle = start;
            this.EndAngle = end;
        }
    }
}