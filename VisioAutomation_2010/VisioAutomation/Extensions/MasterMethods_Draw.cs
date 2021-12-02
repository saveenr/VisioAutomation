using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class MasterMethods_Draw
    {
        public static IVisio.Shape DrawOval(this IVisio.Master master, Core.Rectangle rect)
        {
            return master.Wrap().DrawOval(rect);
        }

        public static IVisio.Shape DrawRectangle(this IVisio.Master master, Core.Rectangle rect)
        {
            return master.Wrap().DrawRectangle(rect);
        }

        public static IVisio.Shape DrawBezier(this IVisio.Master master, IList<Core.Point> points)
        {
            return master.Wrap().DrawBezier(points);
        }

        public static IVisio.Shape DrawPolyline(this IVisio.Master master, IList<Core.Point> points)
        {
            return master.Wrap().DrawPolyline(points);
        }

        public static IVisio.Shape DrawLine(
            this IVisio.Master master,
            Core.Point p0,
            Core.Point p1)
        {
            return master.Wrap().DrawLine(p0, p1);
        }

        public static IVisio.Shape DrawQuarterArc(
            this IVisio.Master master,
            Core.Point p0,
            Core.Point p1,
            IVisio.VisArcSweepFlags flags)
        {
            return master.Wrap().DrawQuarterArc(p0, p1, flags);
        }
    }
}