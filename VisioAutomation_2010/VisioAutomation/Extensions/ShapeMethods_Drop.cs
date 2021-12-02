using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioAutomation.Extensions
{
    public static class ShapeMethods_Drop
    {
        public static IVisio.Shape Drop(
            this IVisio.Shape shape,
            IVisio.Master master,
            Core.Point point)
        {
            return shape.Wrap().Drop(master, point);
        }

        public static short[] DropManyU(
            this IVisio.Shape shape,
            IList<IVisio.Master> masters,
            IEnumerable<Core.Point> points)
        {
            return shape.Wrap().DropManyU(masters, points);
        }
    }
}