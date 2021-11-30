using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class PageMethods_Draw
    {



        public static IVisio.Shape DrawOval(this IVisio.Page page, Core.Rectangle rect)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(page);
            return visobjtarget.DrawOval(rect);
        }

        public static IVisio.Shape DrawRectangle(this IVisio.Page page, Core.Rectangle rect)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(page);
            return visobjtarget.DrawRectangle(rect);
        }
        
        public static IVisio.Shape DrawBezier(this IVisio.Page page, IList<Core.Point> points)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(page);
            return visobjtarget.DrawBezier(points);
        }

        public static IVisio.Shape DrawPolyline(this IVisio.Page page, IList<Core.Point> points)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(page);
            return visobjtarget.DrawPolyline(points);
        }

        public static IVisio.Shape DrawLine(
            this IVisio.Page page,
            Core.Point p0,
            Core.Point p1)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(page);
            return visobjtarget.DrawLine(p0, p1);
        }

        public static IVisio.Shape DrawQuarterArc(
            this IVisio.Page page,
            Core.Point p0,
            Core.Point p1,
            IVisio.VisArcSweepFlags flags)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(page);
            return visobjtarget.DrawQuarterArc(p0, p1, flags);
        }

        public static IVisio.Shape DrawNurbs(
            this IVisio.Page page,
            IList<Core.Point> controlpoints,
            IList<double> knots,
            IList<double> weights,
            int degree)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(page);
            return visobjtarget.DrawNurbs(controlpoints,knots, weights,degree);

        }

    }
}