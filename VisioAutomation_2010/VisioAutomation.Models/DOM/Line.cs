namespace VisioAutomation.Models.Dom
{
    public class Line : BaseShape
    {
        public Drawing.Point P0 { get; private set; }
        public Drawing.Point P1 { get; private set; }

        public Line(double x0, double y0, double x1, double y1)
        {
            this.P0 = new Drawing.Point(x0, y0);
            this.P1 = new Drawing.Point(x1, y1);
        }

        public Line(Drawing.Point p0, Drawing.Point p1)
        {
            this.P0 = p0;
            this.P1 = p1;
        }
    }
}