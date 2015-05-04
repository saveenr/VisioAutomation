using System.Linq;
using VA=VisioAutomation;

namespace VisioAutomation.Drawing
{
    public class BezierCurve
    {
        public Point[] ControlPoints { get; private set; }
        public int Degree { get; set; }

        public BezierCurve(Point[] controlpoints, int degree)
        {
            if (controlpoints== null)
            {
                throw new System.ArgumentNullException("controlpoints");
            }

            if (degree < 1)
            {
                throw new System.ArgumentOutOfRangeException("degree");                
            }

            this.ControlPoints = controlpoints;
            this.Degree = degree;
        }

        public static BezierCurve FromEllipse(Point center, Size radius)
        {
            var pt1 = new Point(0, radius.Height); // top
            var pt2 = new Point(radius.Width, 0); // right
            var pt3 = new Point(0, -radius.Height); // bottom
            var pt4 = new Point(-radius.Width, 0); // left

            double dx = radius.Width * 4.0 * (System.Math.Sqrt(2) - 1) / 3;
            double dy = radius.Height * 4.0 * (System.Math.Sqrt(2) - 1) / 3;

            var curve_ControlPoints = new []
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
            var curve_Degree = 3;
            
            var curve = new BezierCurve(curve_ControlPoints, curve_Degree);
            return curve;
        }
    }
}