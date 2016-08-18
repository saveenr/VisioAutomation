using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ColorsMethods
    {
        public static IEnumerable<Color> ToEnumerable(this Microsoft.Office.Interop.Visio.Colors colors)
        {
            return VisioAutomation.Colors.ColorHelper.ToEnumerable(colors);
        }
    }
}
