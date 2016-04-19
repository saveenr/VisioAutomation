using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ShapeMethods
    {
        public static IVisio.Shape DrawLine(this IVisio.Shape shape, Drawing.Point p1, Drawing.Point p2)
        {
            var surface = new Drawing.DrawingSurface(shape);
            var s = surface.DrawLine(p1, p2);
            return s;
        }

        public static IVisio.Shape DrawQuarterArc(this IVisio.Shape shape, Drawing.Point p0, Drawing.Point p1, IVisio.VisArcSweepFlags flags)
        {
            var surface = new Drawing.DrawingSurface(shape);
            var s = surface.DrawQuarterArc(p0, p1, flags);
            return s;
        }

        public static Drawing.Rectangle GetBoundingBox(this IVisio.Shape shape, IVisio.VisBoundingBoxArgs args)
        {
            var surface = new Drawing.DrawingSurface(shape);
            var r = surface.GetBoundingBox(args);
            return r;
        }

        public static Drawing.Point XYFromPage(this IVisio.Shape shape, Drawing.Point xy)
        {
            // MSDN: http://msdn.microsoft.com/en-us/library/office/ff767213.aspx
            double xprime;
            double yprime;
            shape.XYFromPage(xy.X, xy.Y, out xprime, out yprime);
            return new Drawing.Point(xprime, yprime);
        }

        public static Drawing.Point XYToPage(this IVisio.Shape shape, Drawing.Point xy)
        {
            // MSDN: http://msdn.microsoft.com/en-us/library/office/ff766239.aspx
            double xprime;
            double yprime;
            shape.XYToPage(xy.X, xy.Y, out xprime, out yprime);
            return new Drawing.Point(xprime, yprime);
        }

        public static IEnumerable<IVisio.Shape> AsEnumerable(this IVisio.Shapes shapes)
        {
            int count = shapes.Count;
            for (int i = 0; i < count; i++)
            {
                yield return shapes[i + 1];
            }
        }

        public static IList<IVisio.Shape> GetShapesFromIDs(this IVisio.Shapes shapes, IList<short> shapeids)
        {
            return Shapes.ShapeHelper.GetShapesFromIDs(shapes, shapeids);
        }
    }
}