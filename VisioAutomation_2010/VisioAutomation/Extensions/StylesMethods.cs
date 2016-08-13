using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class StylesMethods
    {
        public static IEnumerable<Style> ToEnumerable(this Microsoft.Office.Interop.Visio.Styles styles)
        {
            return VisioAutomation.Styles.StyleHelper.ToEnumerable(styles);
        }
        
        public static string[] GetNamesU(this Microsoft.Office.Interop.Visio.Styles styles)
        {
            return VisioAutomation.Styles.StyleHelper.GetNamesU(styles);
        }
    }
}