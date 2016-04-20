using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Colors
{
    public static class ColorHelper
    {
        public static IEnumerable<IVisio.Color> ToEnumerable(IVisio.Colors colors)
        {
            int count = colors.Count;
            for (int i = 0; i < count; i++)
            {
                yield return colors[i];
            }
        }
    }
}