using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class SelectionMethods
    {
        public static IEnumerable<IVisio.Shape> ToEnumerable(this IVisio.Selection selection)
        {
            short count16 = selection.Count16;
            for (short i = 0; i < count16; i++)
            {
                yield return selection[i + 1];
            }
        }

        public static IList<IVisio.Shape> ToList(this IVisio.Selection selection)
        {
            int count = selection.Count;
            var list = new List<IVisio.Shape>(count);
            for (int i = 0; i < count; i++)
            {
                list.Add(selection[i + 1]);
            }
            return list;
        }

        public static Geometry.Rectangle GetBoundingBox(this IVisio.Selection selection, IVisio.VisBoundingBoxArgs args)
        {
            double bbx0, bby0, bbx1, bby1;
            selection.BoundingBox((short)args, out bbx0, out bby0, out bbx1, out bby1);
            var r = new Geometry.Rectangle(bbx0, bby0, bbx1, bby1);
            return r;
        }

        public static int[] GetIDs(this IVisio.Selection selection)
        {
            System.Array ids_sa;
            selection.GetIDs(out ids_sa);
            int[] ids = (int[])ids_sa;
            return ids;
        }
    }
}