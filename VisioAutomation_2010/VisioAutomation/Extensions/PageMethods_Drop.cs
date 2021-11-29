using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class PageMethods_Drop
    {
        public static IVisio.Shape Drop(
            this IVisio.Page page,
            IVisio.Master master,
            Core.Point point)
        {
            var output = page.Drop(master, point.X, point.Y);
            return output;
        }

        public static short[] DropManyU(
            this IVisio.Page page,
            IList<IVisio.Master> masters,
            IEnumerable<Core.Point> points)
        {
            return Pages.PageHelper.DropManyU(page, masters, points);
        }
    }
}