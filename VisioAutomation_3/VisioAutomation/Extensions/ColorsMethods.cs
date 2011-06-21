using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Extensions
{
    public static class ColorsMethods
    {
        public static IEnumerable<IVisio.Color> AsEnumerable(this IVisio.Colors colors)
        {
            for (int i = 0; i < colors.Count; i++)
            {
                yield return colors[i];
            }
        }
    }
}