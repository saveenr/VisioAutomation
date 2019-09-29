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
            return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToEnumerable(() => masters.Count,
                i => masters[i + 1]);
        }

        public static List<IVisio.Master> ToList(this IVisio.Masters masters)
        {
            return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToList(() => masters.Count,
                i => masters[i + 1]);
        }

    }
}