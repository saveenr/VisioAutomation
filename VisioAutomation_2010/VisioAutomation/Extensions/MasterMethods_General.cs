using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class MasterMethods_General
    {
        public static Core.Rectangle GetBoundingBox(this IVisio.Master master, IVisio.VisBoundingBoxArgs args)
        {
            double bbx0, bby0, bbx1, bby1;
            master.BoundingBox((short)args, out bbx0, out bby0, out bbx1, out bby1);
            var r = new Core.Rectangle(bbx0, bby0, bbx1, bby1);
            return r;
        }

        public static IEnumerable<IVisio.Master> ToEnumerable(this IVisio.Masters masters)
        {
            return Internal.Extensions.ExtensionHelpers.ToEnumerable(() => masters.Count,
                i => masters[i + 1]);
        }

        public static List<IVisio.Master> ToList(this IVisio.Masters masters)
        {
            return Internal.Extensions.ExtensionHelpers.ToList(() => masters.Count,
                i => masters[i + 1]);
        }

        public static IVisio.Shape Drop(
            this IVisio.Master master1,
            IVisio.Master master2,
            Core.Point point)
        {
            var output = master1.Drop(master2, point.X, point.Y);
            return output;
        }
    }
}