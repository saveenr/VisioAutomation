using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class MasterMethods_Draw
    {
        public static IVisio.Shape DrawOval(this IVisio.Master master, Core.Rectangle rect)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master);
            return visobjtarget.DrawOval(rect);
        }
        
        public static IVisio.Shape DrawRectangle(this IVisio.Master master, Core.Rectangle rect)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master);
            return visobjtarget.DrawRectangle(rect);
        }

        public static IVisio.Shape DrawBezier(this IVisio.Master master, IList<Core.Point> points)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master);
            return visobjtarget.DrawBezier(points);
        }
        
        public static IVisio.Shape DrawPolyline(this IVisio.Master master, IList<Core.Point> points)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master);
            return visobjtarget.DrawPolyline(points);
        }

        public static IVisio.Shape DrawLine(
            this IVisio.Master master,
            Core.Point p0,
            Core.Point p1)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master);
            return visobjtarget.DrawLine(p0, p1);
        }
        
        public static IVisio.Shape DrawQuarterArc(
            this IVisio.Master master,
            Core.Point p0,
            Core.Point p1,
            IVisio.VisArcSweepFlags flags)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master);
            return visobjtarget.DrawQuarterArc(p0, p1, flags);
        }
    }
}