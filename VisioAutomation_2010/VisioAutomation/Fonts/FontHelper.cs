using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Fonts
{
    public static class FontHelper
    {
        public static IEnumerable<Font> ToEnumerable(Microsoft.Office.Interop.Visio.Fonts fonts)
        {
            short count = fonts.Count;
            for (int i = 0; i < count; i++)
            {
                yield return fonts[i + 1];
            }
        }
    }
}