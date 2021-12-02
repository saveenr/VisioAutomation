using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class MasterMethods_Drop
    {
        public static IVisio.Shape Drop(
            this IVisio.Master master1,
            IVisio.Master master2,
            Core.Point point)
        {
            return master1.Wrap().Drop(master2, point);
        }

        public static short[] DropManyU(
            this IVisio.Master master,
            IList<IVisio.Master> masters,
            IEnumerable<Core.Point> points)
        {
            return master.Wrap().DropManyU(masters, points);
        }
    }
}