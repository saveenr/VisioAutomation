using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class PageMethods_Draw
    {
        public static IVisio.Shape DrawOval(this IVisio.Page page, Core.Rectangle rect)
        {
            return page.Wrap().DrawOval(rect);
        }

        public static IVisio.Shape DrawRectangle(this IVisio.Page page, Core.Rectangle rect)
        {
            return page.Wrap().DrawRectangle(rect);
        }

        public static IVisio.Shape DrawBezier(this IVisio.Page page, IList<Core.Point> points)
        {
            return page.Wrap().DrawBezier(points);
        }

        public static IVisio.Shape DrawPolyline(this IVisio.Page page, IList<Core.Point> points)
        {
            return page.Wrap().DrawPolyline(points);
        }

        public static IVisio.Shape DrawLine(
            this IVisio.Page page,
            Core.Point p0,
            Core.Point p1)
        {
            return page.Wrap().DrawLine(p0, p1);
        }

        public static IVisio.Shape DrawQuarterArc(
            this IVisio.Page page,
            Core.Point p0,
            Core.Point p1,
            IVisio.VisArcSweepFlags flags)
        {
            return page.Wrap().DrawQuarterArc(p0, p1, flags);
        }

        public static IVisio.Shape DrawNurbs(
            this IVisio.Page page,
            IList<Core.Point> controlpoints,
            IList<double> knots,
            IList<double> weights,
            int degree)
        {
            return page.Wrap().DrawNurbs(controlpoints, knots, weights, degree);
        }
    }
}