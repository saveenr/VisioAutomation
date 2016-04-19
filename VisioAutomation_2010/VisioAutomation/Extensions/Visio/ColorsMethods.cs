using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ColorsMethods
    {
        public static IEnumerable<IVisio.Color> AsEnumerable(this IVisio.Colors colors)
        {
            int count = colors.Count;
            for (int i = 0; i < count; i++)
            {
                yield return colors[i];
            }
        }
    }
}