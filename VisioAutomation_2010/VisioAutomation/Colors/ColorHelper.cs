using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Colors
{
    public static class ColorHelper
    {
        public static IEnumerable<Color> ToEnumerable(Microsoft.Office.Interop.Visio.Colors colors)
        {
            int count = colors.Count;
            for (int i = 0; i < count; i++)
            {
                yield return colors[i];
            }
        }
    }
}