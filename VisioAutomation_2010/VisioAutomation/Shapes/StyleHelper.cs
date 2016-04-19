using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Styles
{
    public static class StyleHelper
    {
        public static IEnumerable<Style> ToEnumerable(Microsoft.Office.Interop.Visio.Styles styles)
        {
            int count = styles.Count;
            for (int i = 0; i < count; i++)
            {
                yield return styles[i + 1];
            }
        }

        public static string[] GetNamesU(Microsoft.Office.Interop.Visio.Styles styles)
        {
            System.Array names_sa;
            styles.GetNamesU(out names_sa);
            string[] names = (string[])names_sa;
            return names;
        }
    }
}