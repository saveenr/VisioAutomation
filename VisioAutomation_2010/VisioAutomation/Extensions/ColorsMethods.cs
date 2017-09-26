using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ColorsMethods
    {
        public static IEnumerable<IVisio.Color> ToEnumerable(this IVisio.Colors colors)
        {
            int count = colors.Count;
            for (int i = 0; i < count; i++)
            {
                yield return colors[i];
            }
        }

        public static IList<IVisio.Color> ToList(this IVisio.Colors colors)
        {
            int count = colors.Count;
            var list = new List<IVisio.Color>(count);
            for (int i = 0; i < count; i++)
            {
                list.Add(colors[i]);
            }
            return list;
        }
    }
}
