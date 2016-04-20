using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Fonts
{
    public static class FontHelper
    {
        public static IEnumerable<IVisio.Font> ToEnumerable(IVisio.Fonts fonts)
        {
            short count = fonts.Count;
            for (int i = 0; i < count; i++)
            {
                yield return fonts[i + 1];
            }
        }
    }
}