using System.Collections.Generic;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{

    public static class ShapeHelper
    {
        public static IVisio.Shape DrawLine(IVisio.Shape shape, Drawing.Point p1, Drawing.Point p2)
        {
            var surface = new Drawing.DrawingSurface(shape);
            var s = surface.DrawLine(p1, p2);
            return s;
        }

        public static IVisio.Shape DrawQuarterArc(IVisio.Shape shape, Drawing.Point p0, Drawing.Point p1, IVisio.VisArcSweepFlags flags)
        {
            var surface = new Drawing.DrawingSurface(shape);
            var s = surface.DrawQuarterArc(p0, p1, flags);
            return s;
        }

        public static Drawing.Rectangle GetBoundingBox(IVisio.Shape shape, IVisio.VisBoundingBoxArgs args)
        {
            var surface = new Drawing.DrawingSurface(shape);
            var r = surface.GetBoundingBox(args);
            return r;
        }

        public static Drawing.Point XYFromPage(IVisio.Shape shape, Drawing.Point xy)
        {
            // MSDN: http://msdn.microsoft.com/en-us/library/office/ff767213.aspx
            double xprime;
            double yprime;
            shape.XYFromPage(xy.X, xy.Y, out xprime, out yprime);
            return new Drawing.Point(xprime, yprime);
        }

        public static Drawing.Point XYToPage(IVisio.Shape shape, Drawing.Point xy)
        {
            // MSDN: http://msdn.microsoft.com/en-us/library/office/ff766239.aspx
            double xprime;
            double yprime;
            shape.XYToPage(xy.X, xy.Y, out xprime, out yprime);
            return new Drawing.Point(xprime, yprime);
        }

        public static IEnumerable<IVisio.Shape> ToEnumerable(IVisio.Shapes shapes)
        {
            int count = shapes.Count;
            for (int i = 0; i < count; i++)
            {
                yield return shapes[i + 1];
            }
        }

        /// <summary>
        /// Enumerates all shapes contained by a set of shapes recursively
        /// </summary>
        /// <param name="shapes">the set of shapes to start the enumeration</param>
        /// <returns>The enumeration</returns>
        public static IList<IVisio.Shape> GetNestedShapes(IEnumerable<IVisio.Shape> shapes)
        {
            if (shapes == null)
            {
                throw new System.ArgumentNullException(nameof(shapes));
            }

            var result = new List<IVisio.Shape>();
            var stack = new Stack<IVisio.Shape>(shapes);

            while (stack.Count > 0)
            {
                var s = stack.Pop();
                var subshapes = s.Shapes;
                if (subshapes.Count > 0)
                {
                    foreach (var child in subshapes.ToEnumerable())
                    {
                        stack.Push(child);
                    }
                }

                result.Add(s);
            }

            return result;
        }

        public static IList<IVisio.Shape> GetNestedShapes(IVisio.Shape shape)
        {
            if (shape== null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            var shapes = new[] {shape};

            return ShapeHelper.GetNestedShapes(shapes);
        }

        public static IList<IVisio.Shape> GetShapesFromIDs(IVisio.Shapes shapes, IList<short> shapeids)
        {
            var shape_objs = new List<IVisio.Shape>(shapeids.Count);
            foreach (short shapeid in shapeids)
            {
                var shape = shapes.ItemFromID16[shapeid];
                shape_objs.Add(shape);
            }
            return shape_objs;
        }
    }
}