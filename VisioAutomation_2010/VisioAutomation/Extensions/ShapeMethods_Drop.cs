using System.Collections.Generic;
using System.Linq;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ShapeMethods_Drop
    {

        public static IVisio.Shape Drop(
            this IVisio.Shape shape,
            IVisio.Master master,
            Core.Point point)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(shape);
            return visobjtarget.Drop(master, point);
        }

        public static short[] DropManyU(
            this IVisio.Shape shape,
            IList<IVisio.Master> masters,
            IEnumerable<Core.Point> points)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(shape);
            return visobjtarget.DropManyU(masters, points);
        }
    }
}