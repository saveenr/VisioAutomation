
namespace VisioAutomation.Extensions
{
    public static class SelectionMethods
    {
        public static IEnumerable<IVisio.Shape> ToEnumerable(this IVisio.Selection selection)
        {
            return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToEnumerable(() => selection.Count, i => selection[i + 1]);
        }

        public static List<IVisio.Shape> ToList(this IVisio.Selection selection)
        {
            return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToList(() => selection.Count, i => selection[i + 1]);
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