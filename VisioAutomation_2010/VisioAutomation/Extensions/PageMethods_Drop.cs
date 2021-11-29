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
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(page);
            return VisioAutomation.Internal.Extensions.ExtensionHelpers._Drop(visobjtarget, master, point);
        }

        public static short[] DropManyU(
            this IVisio.Page page,
            IList<IVisio.Master> masters,
            IEnumerable<Core.Point> points)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(page);
            return VisioAutomation.Internal.Extensions.ExtensionHelpers._DropManyU(visobjtarget, masters, points);
        }
    }
}