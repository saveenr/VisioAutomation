using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ColorsMethods
    {
        public static IEnumerable<IVisio.Color> ToEnumerable(this IVisio.Colors colors)
        {
            return Extensions.ExtensionHelpers.ToEnumerable(() => colors.Count, i => colors[i]);
        }

        public static List<IVisio.Color> ToList(this IVisio.Colors colors)
        {
            return Extensions.ExtensionHelpers.ToList(() => colors.Count, i => colors[i]);
        }
    }
}
