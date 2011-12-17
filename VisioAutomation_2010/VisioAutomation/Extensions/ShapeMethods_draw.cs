using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Extensions
{
    public static partial class ShapeMethods
    {
        public static IVisio.Shape DrawLine(this IVisio.Shape shape, VA.Drawing.Point p1, VA.Drawing.Point p2)
        {
            var s = shape.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
            return s;
        }

        public static IVisio.Shape DrawQuarterArc(this IVisio.Shape shape, VA.Drawing.Point p0, VA.Drawing.Point p1, IVisio.VisArcSweepFlags flags)
        {
            return shape.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
        }
    }
}