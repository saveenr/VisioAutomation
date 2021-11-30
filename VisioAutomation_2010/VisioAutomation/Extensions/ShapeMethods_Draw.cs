using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ShapeMethods_Draw
    {
        public static IVisio.Shape DrawOval(this IVisio.Shape shape, Core.Rectangle rect)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(shape);
            return visobjtarget.DrawOval(rect);
        }

        public static IVisio.Shape DrawRectangle(this IVisio.Shape shape, Core.Rectangle rect)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(shape);
            return visobjtarget.DrawRectangle(rect);
        }

        public static IVisio.Shape DrawBezier(this IVisio.Shape  shape, IList<Core.Point> points)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(shape);
            return visobjtarget.DrawBezier(points);
        }


        public static IVisio.Shape DrawPolyline(this IVisio.Shape shape, IList<Core.Point> points)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(shape);
            return visobjtarget.DrawPolyline(points);
        }

        public static IVisio.Shape DrawLine(
            this IVisio.Shape shape,
            Core.Point p0,
            Core.Point p1)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(shape);
            return visobjtarget.DrawLine(p0, p1);
        }

        public static IVisio.Shape DrawQuarterArc(
            this IVisio.Shape shape,
            Core.Point p0,
            Core.Point p1,
            IVisio.VisArcSweepFlags flags)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(shape);
            return visobjtarget.DrawQuarterArc(p0, p1, flags);
        }
        
        public static IVisio.Shape DrawNurbs(
            this IVisio.Shape shape,
            IList<Core.Point> controlpoints,
            IList<double> knots,
            IList<double> weights,
            int degree)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(shape);
            return visobjtarget.DrawNurbs(controlpoints, knots, weights, degree);

        }

    }
}