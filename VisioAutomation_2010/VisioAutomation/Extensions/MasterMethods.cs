using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class MasterMethods
    {
        public static Geometry.Rectangle GetBoundingBox(this IVisio.Master master, IVisio.VisBoundingBoxArgs args)
        {
            var surface = new VisioAutomation.SurfaceTarget(master);
            return surface.GetBoundingBox(args);
        }

        public static IEnumerable<IVisio.Master> ToEnumerable(this IVisio.Masters masters)
        {
            return ExtensionHelpers.ToEnumerable(() => masters.Count, i => masters[i + 1]);
        }

        public static List<IVisio.Master> ToList(this IVisio.Masters masters)
        {
            return ExtensionHelpers.ToList(() => masters.Count, i => masters[i + 1]);
        }

        public static string[] GetNamesU(this IVisio.Masters masters)
        {
            System.Array names_sa;
            masters.GetNamesU(out names_sa);
            string[] names = (string[])names_sa;
            return names;
        }
    }
}