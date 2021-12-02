using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ShapeMethods_Draw
    {
        public static IVisio.Shape DrawOval(this IVisio.Shape shape, Core.Rectangle rect)
        {
            return shape.Wrap().DrawOval(rect);
        }

        public static IVisio.Shape DrawRectangle(this IVisio.Shape shape, Core.Rectangle rect)
        {
            return shape.Wrap().DrawRectangle(rect);
        }

        public static IVisio.Shape DrawBezier(this IVisio.Shape shape, IList<Core.Point> points)
        {
            return shape.Wrap().DrawBezier(points);
        }


        public static IVisio.Shape DrawPolyline(this IVisio.Shape shape, IList<Core.Point> points)
        {
            return shape.Wrap().DrawPolyline(points);
        }

        public static IVisio.Shape DrawLine(
            this IVisio.Shape shape,
            Core.Point p0,
            Core.Point p1)
        {
            return shape.Wrap().DrawLine(p0, p1);
        }

        public static IVisio.Shape DrawQuarterArc(
            this IVisio.Shape shape,
            Core.Point p0,
            Core.Point p1,
            IVisio.VisArcSweepFlags flags)
        {
            return shape.Wrap().DrawQuarterArc(p0, p1, flags);
        }

        public static IVisio.Shape DrawNurbs(
            this IVisio.Shape shape,
            IList<Core.Point> controlpoints,
            IList<double> knots,
            IList<double> weights,
            int degree)
        {
            return shape.Wrap().DrawNurbs(controlpoints, knots, weights, degree);
        }
    }
}