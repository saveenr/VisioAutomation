using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class PageMethods_Draw
    {



        public static Microsoft.Office.Interop.Visio.Shape DrawOval(this Microsoft.Office.Interop.Visio.Page page, Core.Rectangle rect)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(page);
            return visobjtarget.DrawOval(rect);
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawRectangle(this Microsoft.Office.Interop.Visio.Page page, Core.Rectangle rect)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(page);
            return visobjtarget.DrawRectangle(rect);
        }
        
        public static Microsoft.Office.Interop.Visio.Shape DrawBezier(this IVisio.Page page, IList<Core.Point> points)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(page);
            return visobjtarget.DrawBezier(points);
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawPolyline(this IVisio.Page page, IList<Core.Point> points)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(page);
            return visobjtarget.DrawPolyline(points);
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawLine(
            this Microsoft.Office.Interop.Visio.Page page,
            Core.Point p0,
            Core.Point p1)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(page);
            return visobjtarget.DrawLine(p0, p1);
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawQuarterArc(
            this Microsoft.Office.Interop.Visio.Page page,
            Core.Point p0,
            Core.Point p1,
            Microsoft.Office.Interop.Visio.VisArcSweepFlags flags)
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