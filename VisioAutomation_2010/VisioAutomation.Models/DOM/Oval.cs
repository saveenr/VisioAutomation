namespace VisioAutomation.Models.Dom
{
    public class Oval : BaseShape
    {
        public VisioAutomation.Core.Point P0 { get; }
        public VisioAutomation.Core.Point P1 { get; }

        public Oval(double x0, double y0, double x1, double y1)
        {
            this.P0 = new VisioAutomation.Core.Point(x0, y0);
            this.P1 = new VisioAutomation.Core.Point(x1, y1);
        }

        public Oval(VisioAutomation.Core.Point p0, VisioAutomation.Core.Point p1)
        {
            this.P0 = p0;
            this.P1 = p1;
        }

        public Oval(VisioAutomation.Core.Rectangle r0)
        {
            this.P0 = r0.LowerLeft;
            this.P1 = r0.UpperRight;
        }
    }
}
