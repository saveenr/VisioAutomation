using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class MasterMethods_Draw
    {
        public static Microsoft.Office.Interop.Visio.Shape DrawOval(this Microsoft.Office.Interop.Visio.Master master, Core.Rectangle rect)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master);
            return visobjtarget.DrawOval(rect);
        }
        
        public static Microsoft.Office.Interop.Visio.Shape DrawRectangle(this Microsoft.Office.Interop.Visio.Master master, Core.Rectangle rect)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master);
            return visobjtarget.DrawRectangle(rect);
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawBezier(this IVisio.Master master, IList<Core.Point> points)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master);
            return visobjtarget.DrawBezier(points);
        }
        
        public static Microsoft.Office.Interop.Visio.Shape DrawPolyline(this Microsoft.Office.Interop.Visio.Master master, IList<Core.Point> points)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master);
            return visobjtarget.DrawPolyline(points);
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawLine(
            this Microsoft.Office.Interop.Visio.Master master,
            Core.Point p0,
            Core.Point p1)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master);
            return visobjtarget.DrawLine(p0, p1);
        }
        
        public static Microsoft.Office.Interop.Visio.Shape DrawQuarterArc(
            this Microsoft.Office.Interop.Visio.Master master,
            Core.Point p0,
            Core.Point p1,
            Microsoft.Office.Interop.Visio.VisArcSweepFlags flags)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master);
            return visobjtarget.DrawQuarterArc(p0, p1, flags);
        }
    }
}