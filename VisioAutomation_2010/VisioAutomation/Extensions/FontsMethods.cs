using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class FontsMethods
    {
        public static IEnumerable<Font> ToEnumerable(this Microsoft.Office.Interop.Visio.Fonts fonts)
        {
            return VisioAutomation.Fonts.FontHelper.ToEnumerable(fonts);
        }
    }
}