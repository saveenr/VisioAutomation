using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Extensions
{
    public static class ShapesMethods
    {
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
            return VA.ShapeHelper.GetShapesFromIDs(shapes, shapeids);
        }

    }
}