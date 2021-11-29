using System.Collections.Generic;

namespace VisioAutomation.Extensions
{
    public static class PageMethods_Drop
    {
        public static Microsoft.Office.Interop.Visio.Shape Drop(
            this Microsoft.Office.Interop.Visio.Page page,
            Microsoft.Office.Interop.Visio.Master master,
            Core.Point point)
        {
            var output = page.Drop(master, point.X, point.Y);
            return output;
        }

        public static short[] DropManyU(
            this Microsoft.Office.Interop.Visio.Page page,
            IList<Microsoft.Office.Interop.Visio.Master> masters,
            IEnumerable<Core.Point> points)
        {
            return Pages.PageHelper.DropManyU(page, masters, points);
        }
    }
}