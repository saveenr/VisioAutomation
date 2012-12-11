using System.Linq;
using VA=VisioAutomation;

namespace VisioAutomation.Drawing
{
    public class BezierCurve
    {
        public VA.Drawing.Point[] ControlPoints { get; private set; }
        public int Degree { get; set; }

        public BezierCurve() :
            this(null, -1)
        {
        }

        public BezierCurve(VA.Drawing.Point[] controlpoints, int degree)
        {
            this.ControlPoints = controlpoints;
            this.Degree = degree;
        }

        public static BezierCurve FromEllipse(VA.Drawing.Point center, VA.Drawing.Size radius)
        {
            var curve = new BezierCurve();

            var pt1 = new VA.Drawing.Point(0, radius.Height); // top
            var pt2 = new VA.Drawing.Point(radius.Width, 0); // right
            var pt3 = new VA.Drawing.Point(0, -radius.Height); // bottom
            var pt4 = new VA.Drawing.Point(-radius.Width, 0); // left

            double dx = radius.Width * 4.0 * (System.Math.Sqrt(2) - 1) / 3;
            double dy = radius.Height * 4.0 * (System.Math.Sqrt(2) - 1) / 3;

            curve.ControlPoints = new []
                                      {
                                          pt1,
                                          pt1.Add(dx, 0),
                                          pt2.Add(0, dy),
                                          pt2,
                                          pt2.Add(0, -dy),
                                          pt3.Add(dx, 0),
                                          pt3,
                                          pt3.Add(-dx, 0),
                                          pt4.Add(0, -dy),
                                          pt4,
                                          pt4.Add(0, dy),
                                          pt1.Add(-dx, 0),
                                          pt1
                                      }
                .Select(p => p + center).ToArray();

            curve.Degree = 3;

            return curve;
        }
    }
}