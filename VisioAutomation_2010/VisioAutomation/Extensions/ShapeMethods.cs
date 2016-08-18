using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{

    public static class ShapeMethods
    {
        public static IVisio.Shape DrawLine(this IVisio.Shape shape, Drawing.Point p1, Drawing.Point p2)
        {
            return Shapes.ShapeHelper.DrawLine(shape, p1, p2);
        }

        public static IVisio.Shape DrawQuarterArc(this IVisio.Shape shape, Drawing.Point p0, Drawing.Point p1, IVisio.VisArcSweepFlags flags)
        {
            return Shapes.ShapeHelper.DrawQuarterArc(shape, p0, p1, flags);
        }

        public static Drawing.Rectangle GetBoundingBox(this IVisio.Shape shape, IVisio.VisBoundingBoxArgs args)
        {
            return Shapes.ShapeHelper.GetBoundingBox(shape, args);
        }

        public static Drawing.Point XYFromPage(this IVisio.Shape shape, Drawing.Point xy)
        {
            return Shapes.ShapeHelper.XYFromPage(shape, xy);
        }

        public static Drawing.Point XYToPage(this IVisio.Shape shape, Drawing.Point xy)
        {
            return Shapes.ShapeHelper.XYToPage(shape, xy);
        }

        public static IEnumerable<IVisio.Shape> ToEnumerable(this IVisio.Shapes shapes)
        {
            return Shapes.ShapeHelper.ToEnumerable(shapes);
        }

        public static IList<IVisio.Shape> GetShapesFromIDs(this IVisio.Shapes shapes, IList<short> shapeids)
        {
            return Shapes.ShapeHelper.GetShapesFromIDs(shapes, shapeids);
        }
    }
}