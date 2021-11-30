using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ShapeMethods_Draw
    {
        public static Microsoft.Office.Interop.Visio.Shape DrawOval(this Microsoft.Office.Interop.Visio.Shape shape, Core.Rectangle rect)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(shape);
            return visobjtarget.DrawOval(rect);
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawRectangle(this Microsoft.Office.Interop.Visio.Shape shape, Core.Rectangle rect)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(shape);
            return visobjtarget.DrawRectangle(rect);
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawBezier(this IVisio.Shape  shape, IList<Core.Point> points)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(shape);
            return visobjtarget.DrawBezier(points);
        }


        public static Microsoft.Office.Interop.Visio.Shape DrawPolyline(this Microsoft.Office.Interop.Visio.Shape shape, IList<Core.Point> points)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(shape);
            return visobjtarget.DrawPolyline(points);
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawLine(
            this Microsoft.Office.Interop.Visio.Shape shape,
            Core.Point p0,
            Core.Point p1)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(shape);
            return visobjtarget.DrawLine(p0, p1);
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawQuarterArc(
            this Microsoft.Office.Interop.Visio.Shape shape,
            Core.Point p0,
            Core.Point p1,
            Microsoft.Office.Interop.Visio.VisArcSweepFlags flags)
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