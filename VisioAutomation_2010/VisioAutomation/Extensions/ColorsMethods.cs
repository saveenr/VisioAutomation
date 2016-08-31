using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ColorsMethods
    {
        public static IEnumerable<IVisio.Color> ToEnumerable(this IVisio.Colors colors)
        {
            return VisioAutomation.Colors.ColorHelper.ToEnumerable(colors);
        }
    }
}
