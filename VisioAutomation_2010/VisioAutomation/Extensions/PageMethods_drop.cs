using System.Collections.Generic;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static partial class PageMethods
    {
        public static IVisio.Shape Drop(
            this IVisio.Page page, 
            IVisio.Master master,
            VA.Drawing.Point point)
        {
            if (master == null)
            {
                throw new System.ArgumentNullException("master");
            }

            return page.Drop(master, point.X, point.Y);
        }

        public static short[] DropManyU(
            this IVisio.Page page, 
            IList<IVisio.Master> masters,
            IEnumerable<VA.Drawing.Point> points)
        {
            short[] shapeids = VA.Pages.PageHelper.DropManyU(page, masters, points);
            return shapeids;
        }
    }
}