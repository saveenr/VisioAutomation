using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class FontsMethods
    {
        public static IEnumerable<IVisio.Font> ToEnumerable(this IVisio.Fonts fonts)
        {
            return Internal.Extensions.ExtensionHelpers.ToEnumerable(() => fonts.Count, i => fonts[i + 1]);
        }

        public static List<IVisio.Font> ToList(this IVisio.Fonts fonts)
        {
            return Internal.Extensions.ExtensionHelpers.ToList(() => fonts.Count, i => fonts[i + 1]);
        }
    }
}