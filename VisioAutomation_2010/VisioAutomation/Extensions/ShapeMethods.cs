using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ShapeMethods
    {
        public static IVisio.Shape DrawLine(
            this IVisio.Shape shape,
            Geometry.Point p1, Geometry.Point p2)
        {
            var surface = new SurfaceTarget(shape);
            var s = surface.DrawLine(p1, p2);
            return s;
        }

        public static IVisio.Shape DrawQuarterArc(
            this IVisio.Shape shape,
            Geometry.Point p0,
            Geometry.Point p1, 
            IVisio.VisArcSweepFlags flags)
        {
            var surface = new SurfaceTarget(shape);
            var s = surface.DrawQuarterArc(p0, p1, flags);
            return s;
        }

        public static Geometry.Rectangle GetBoundingBox(
            this IVisio.Shape shape, 
            IVisio.VisBoundingBoxArgs args)
        {
            var surface = new SurfaceTarget(shape);
            var r = surface.GetBoundingBox(args);
            return r;
        }

        public static Geometry.Point XYFromPage(
            this IVisio.Shape shape,
            Geometry.Point xy)
        {
            // MSDN: http://msdn.microsoft.com/en-us/library/office/ff767213.aspx
            double xprime;
            double yprime;
            shape.XYFromPage(xy.X, xy.Y, out xprime, out yprime);
            return new Geometry.Point(xprime, yprime);
        }

        public static Geometry.Point XYToPage(
            this IVisio.Shape shape,
            Geometry.Point xy)
        {
            // MSDN: http://msdn.microsoft.com/en-us/library/office/ff766239.aspx
            double xprime;
            double yprime;
            shape.XYToPage(xy.X, xy.Y, out xprime, out yprime);
            return new Geometry.Point(xprime, yprime);
        }

        public static IEnumerable<IVisio.Shape> ToEnumerable(this IVisio.Shapes shapes)
        {
            int count = shapes.Count;
            for (int i = 0; i < count; i++)
            {
                yield return shapes[i + 1];
            }
        }

        public static List<IVisio.Shape> ToList(this IVisio.Shapes shapes)
        {
            int count = shapes.Count;
            var list = new List<IVisio.Shape>(count);
            for (int i = 0; i < count; i++)
            {
                list.Add(shapes[i + 1]);
            }
            return list;
        }
    }
}