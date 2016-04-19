using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class FontsMethods
    {
        public static IEnumerable<IVisio.Font> AsEnumerable(this IVisio.Fonts fonts)
        {
            short count = fonts.Count;
            for (int i = 0; i < count; i++)
            {
                yield return fonts[i + 1];
            }
        }
    }
}