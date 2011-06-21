using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;
using VA = VisioAutomation;

namespace VisioAutomation.Extensions
{
    public static class ShapesMethods
    {
        public static IEnumerable<IVisio.Shape> AsEnumerable(this IVisio.Shapes shapes)
        {
            for (int i = 0; i < shapes.Count; i++)
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