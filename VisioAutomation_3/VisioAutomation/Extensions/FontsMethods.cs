using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Extensions
{
    public static class FontsMethods
    {
        public static IEnumerable<IVisio.Font> AsEnumerable(this IVisio.Fonts fonts)
        {
            return fonts.Cast<IVisio.Font>();
        }
    }
}