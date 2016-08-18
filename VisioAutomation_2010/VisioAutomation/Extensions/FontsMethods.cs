using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class FontsMethods
    {
        public static IEnumerable<IVisio.Font> ToEnumerable(this IVisio.Fonts fonts)
        {
            return VisioAutomation.Fonts.FontHelper.ToEnumerable(fonts);
        }
    }
}