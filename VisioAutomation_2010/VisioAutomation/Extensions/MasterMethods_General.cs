using System.Collections.Generic;
using VisioAutomation.Internal;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class MasterMethods_General
    {
        public static Core.Rectangle GetBoundingBox(this IVisio.Master master, IVisio.VisBoundingBoxArgs args)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master);
            return visobjtarget.GetBoundingBox(args);
        }

        public static IEnumerable<IVisio.Master> ToEnumerable(this IVisio.Masters masters)
        {
            return CollectionHelpers.ToEnumerable(() => masters.Count,
                i => masters[i + 1]);
        }

        public static List<IVisio.Master> ToList(this IVisio.Masters masters)
        {
            return CollectionHelpers.ToList(() => masters.Count,
                i => masters[i + 1]);
        }

    }
}