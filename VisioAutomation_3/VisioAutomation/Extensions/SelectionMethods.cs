using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;

namespace VisioAutomation.Extensions
{
    public static class SelectionMethods
    {
        public static VA.Drawing.Rectangle GetBoundingBox(this IVisio.Selection selection, IVisio.VisBoundingBoxArgs args)
        {
            double bbx0, bby0, bbx1, bby1;
            selection.BoundingBox((short) args, out bbx0, out bby0, out bbx1, out bby1);
            var r = new VA.Drawing.Rectangle(bbx0, bby0, bbx1, bby1);
            return r;
        }

        public static IEnumerable<IVisio.Shape> AsEnumerable(this IVisio.Selection selection)
        {
            return selection.Cast<IVisio.Shape>();
        }
        
        public static int[] GetIDs(this IVisio.Selection selection)
        {
            System.Array ids_sa;
            selection.GetIDs(out ids_sa);
            int[] ids = (int[]) ids_sa;
            return ids;
        }
    }
}