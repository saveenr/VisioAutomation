using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class FontsMethods
    {
        public static IEnumerable<IVisio.Font> ToEnumerable(this IVisio.Fonts fonts)
        {
            short count = fonts.Count;
            for (int i = 0; i < count; i++)
            {
                yield return fonts[i + 1];
            }
        }

        public static IList<IVisio.Font> ToList(this IVisio.Fonts fonts)
        {
            int count = fonts.Count;
            var list = new List<IVisio.Font>(count);
            for (int i = 0; i < count; i++)
            {
                list.Add(fonts[i+1]);
            }
            return list;
        }
    }
}