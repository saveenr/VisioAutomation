using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class PageMethods_Draw
    {
        public static Microsoft.Office.Interop.Visio.Shape DrawLine(this Microsoft.Office.Interop.Visio.Page page, Core.Point p1, Core.Point p2)
        {
            var shape = page.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
            return shape;
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawOval(this Microsoft.Office.Interop.Visio.Page page, Core.Rectangle rect)
        {
            var shape = page.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top);
            return shape;
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawRectangle(this Microsoft.Office.Interop.Visio.Page page, Core.Rectangle rect)
        {
            var shape = page.DrawRectangle(rect.Left, rect.Bottom, rect.Right, rect.Top);
            return shape;
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawBezier(this Microsoft.Office.Interop.Visio.Page page, IList<Core.Point> points)
        {
            var doubles_array = VisioAutomation.Core.Point.ToDoubles(points).ToArray();
            short degree = 3;
            short flags = 0;
            var shape = page.DrawBezier(doubles_array, degree, flags);
            return shape;
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawPolyline(this Microsoft.Office.Interop.Visio.Page page, IList<Core.Point> points)
        {
            var shape = page.DrawBezier(points);
            return shape;
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawPolyLine(this Microsoft.Office.Interop.Visio.Page page, IList<Core.Point> points)
        {
            var doubles_array = Core.Point.ToDoubles(points).ToArray();
            var shape = page.DrawPolyline(doubles_array, 0);
            return shape;
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawQuarterArc(
            this Microsoft.Office.Interop.Visio.Page page,
            Core.Point p0,
            Core.Point p1,
            Microsoft.Office.Interop.Visio.VisArcSweepFlags flags)
        {
            var s = page.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            return s;
        }
        
        public static IVisio.Shape DrawNurbs(
            this IVisio.Page page,
            IList<Core.Point> controlpoints,
            IList<double> knots,
            IList<double> weights,
            int degree)
        {

            // flags:
            // None = 0,
            // IVisio.VisDrawSplineFlags.visSpline1D

            var flags = 0;
            double[] pts_dbl_a = Core.Point.ToDoubles(controlpoints).ToArray();
            double[] kts_dbl_a = knots.ToArray();
            double[] weights_dbl_a = weights.ToArray();

            var shape = page.DrawNURBS((short) degree, (short) flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
            return shape;
        }

    }
}