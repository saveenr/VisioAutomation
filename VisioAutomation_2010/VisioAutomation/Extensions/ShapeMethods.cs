using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{

    public static class ShapeMethods
    {
        public static Microsoft.Office.Interop.Visio.Shape DrawLine(this Microsoft.Office.Interop.Visio.Shape shape, Drawing.Point p1, Drawing.Point p2)
        {
            return Shapes.ShapeHelper.DrawLine(shape, p1, p2);
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawQuarterArc(this Microsoft.Office.Interop.Visio.Shape shape, Drawing.Point p0, Drawing.Point p1, Microsoft.Office.Interop.Visio.VisArcSweepFlags flags)
        {
            return Shapes.ShapeHelper.DrawQuarterArc(shape, p0, p1, flags);
        }

        public static Drawing.Rectangle GetBoundingBox(this Microsoft.Office.Interop.Visio.Shape shape, Microsoft.Office.Interop.Visio.VisBoundingBoxArgs args)
        {
            return Shapes.ShapeHelper.GetBoundingBox(shape, args);
        }

        public static Drawing.Point XYFromPage(this Microsoft.Office.Interop.Visio.Shape shape, Drawing.Point xy)
        {
            return Shapes.ShapeHelper.XYFromPage(shape, xy);
        }

        public static Drawing.Point XYToPage(this Microsoft.Office.Interop.Visio.Shape shape, Drawing.Point xy)
        {
            return Shapes.ShapeHelper.XYToPage(shape, xy);
        }

        public static IEnumerable<Shape> ToEnumerable(this Microsoft.Office.Interop.Visio.Shapes shapes)
        {
            return Shapes.ShapeHelper.ToEnumerable(shapes);
        }

        public static IList<Shape> GetShapesFromIDs(this Microsoft.Office.Interop.Visio.Shapes shapes, IList<short> shapeids)
        {
            return Shapes.ShapeHelper.GetShapesFromIDs(shapes, shapeids);
        }
    }
}