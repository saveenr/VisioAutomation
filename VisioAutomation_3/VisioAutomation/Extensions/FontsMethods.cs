using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Extensions
{
    public static class FontsMethods
    {
        public static IEnumerable<IVisio.Font> AsEnumerable(this IVisio.Fonts fonts)
        {
            for (int i = 0; i < fonts.Count; i++)
            {
                yield return fonts[i + 1];
            }
        }
    }
}