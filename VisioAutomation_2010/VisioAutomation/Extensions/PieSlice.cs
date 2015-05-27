namespace VisioAutomation.DOM
{
    public class PieSlice: BaseShape
    {
        public Drawing.Point Center { get; private set; }
        public double Radius { get; private set; }
        public double Start { get; private set; }
        public double End { get; private set; }

        public PieSlice(double x0, double y0, double r, double start, double end)
        {
            this.Center = new Drawing.Point(x0, y0);
            this.Radius= r;
            this.Start = start;
            this.End = end;
        }

        public PieSlice(Drawing.Point p0, double r, double start, double end)
        {
            this.Center = p0;
            this.Radius = r;
            this.Start = start;
            this.End = end;
        }
    }
}