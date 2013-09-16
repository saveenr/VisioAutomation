using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Extensions
{
    public static class ShapeMethods
    {
        public static IVisio.Shape DrawLine(this IVisio.Shape shape, VA.Drawing.Point p1, VA.Drawing.Point p2)
        {
            // MSDN: http://msdn.microsoft.com/en-us/library/office/ff766239.aspx
            var s = shape.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
            return s;
        }

        public static IVisio.Shape DrawQuarterArc(this IVisio.Shape shape, VA.Drawing.Point p0, VA.Drawing.Point p1, IVisio.VisArcSweepFlags flags)
        {
            // MSDN: http://msdn.microsoft.com/en-us/library/office/ff767062(v=office.14).aspx
            return shape.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
        }

        public static VA.Drawing.Rectangle GetBoundingBox(this IVisio.Shape shape, IVisio.VisBoundingBoxArgs args)
        {
            // MSDN: http://msdn.microsoft.com/library/default.asp?url=/library/en-us/vissdk11/html/vimthBoundingBox_HV81900422.asp
            double bbx0, bby0, bbx1, bby1;
            shape.BoundingBox((short)args, out bbx0, out bby0, out bbx1, out bby1);
            var r = new VA.Drawing.Rectangle(bbx0, bby0, bbx1, bby1);
            return r;
        }

        public static VA.Drawing.Point XYFromPage(this IVisio.Shape shape, VA.Drawing.Point xy)
        {
            // MSDN: http://msdn.microsoft.com/en-us/library/office/ff767213.aspx
            double xprime;
            double yprime;
            shape.XYFromPage(xy.X, xy.Y, out xprime, out yprime);
            return new VA.Drawing.Point(xprime, yprime);
        }

        public static VA.Drawing.Point XYToPage(this IVisio.Shape shape, VA.Drawing.Point xy)
        {
            // MSDN: http://msdn.microsoft.com/en-us/library/office/ff766239.aspx
            double xprime;
            double yprime;
            shape.XYToPage(xy.X, xy.Y, out xprime, out yprime);
            return new VA.Drawing.Point(xprime, yprime);
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
            return VA.Shapes.ShapeHelper.GetShapesFromIDs(shapes, shapeids);
        }
    }
}