using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class StylesMethods
    {
        public static IEnumerable<IVisio.Style> AsEnumerable(this IVisio.Styles styles)
        {
            int count = styles.Count;
            for (int i = 0; i < count; i++)
            {
                yield return styles[i + 1];
            }
        }
        
        public static string[] GetNamesU(this IVisio.Styles styles)
        {
            System.Array names_sa;
            styles.GetNamesU(out names_sa);
            string[] names = (string[])names_sa;
            return names;
        }
    }
}