using System.Collections.Generic;
using VisioAutomation.Internal;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ColorsMethods
    {
        public static IEnumerable<IVisio.Color> ToEnumerable(this IVisio.Colors colors)
        {
            return CollectionHelpers.ToEnumerable(() => colors.Count, i => colors[i]);
        }

        public static List<IVisio.Color> ToList(this IVisio.Colors colors)
        {
            return CollectionHelpers.ToList(() => colors.Count, i => colors[i]);
        }
    }
}
