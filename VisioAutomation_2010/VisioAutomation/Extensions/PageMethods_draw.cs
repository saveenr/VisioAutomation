using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.Extensions
{
    public static partial class PageMethods
    {
        public static IVisio.Shape DrawLine(this IVisio.Page page, VA.Drawing.Point p1, VA.Drawing.Point p2)
        {
            var shape = page.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
            return shape;
        }

        public static IVisio.Shape DrawOval(this IVisio.Page page, VA.Drawing.Rectangle rect)
        {
            var shape = page.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top);
            return shape;
        }

        public static IVisio.Shape DrawRectangle(this IVisio.Page page, VA.Drawing.Rectangle rect)
        {
            var shape = page.DrawRectangle(rect.Left, rect.Bottom, rect.Right, rect.Top);
            return shape;
        }

        public static IVisio.Shape DrawBezier(this IVisio.Page page, IList<VA.Drawing.Point> points)
        {
            var doubles_array = VA.Drawing.Point.ToDoubles(points).ToArray();
            var shape = page.DrawBezier(doubles_array, 3, 0);
            return shape;
        }

        public static IVisio.Shape DrawPolyline(this IVisio.Page page, IList<VA.Drawing.Point> points)
        {
            var doubles_array = VA.Drawing.Point.ToDoubles(points).ToArray();
            var shape = page.DrawPolyline(doubles_array, 0);
            return shape;
        }

        public static IVisio.Shape DrawNURBS(this IVisio.Page page, IList<VA.Drawing.Point> controlpoints,
                            IList<double> knots,
                            IList<double> weights, int degree)
        {

            // flags:
            // None = 0,
            // IVisio.VisDrawSplineFlags.visSpline1D

            var flags = 0;
            double[] pts_dbl_a = VA.Drawing.Point.ToDoubles(controlpoints).ToArray();
                double[] kts_dbl_a = knots.ToArray();
                double[] weights_dbl_a = weights.ToArray();

                var shape = page.DrawNURBS((short)degree, (short)flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
                return shape;
        }
    }
}