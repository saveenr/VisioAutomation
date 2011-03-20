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
            return shapes.Cast<IVisio.Shape>();
        }

        public static IList<IVisio.Shape> GetShapesFromIDs(this IVisio.Shapes shapes, IList<short> shapeids)
        {
            return VA.ShapeHelper.GetShapesFromIDs(shapes, shapeids);
        }

    }
}