using System.Collections.Generic;
using VisioAutomation.Internal;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class StylesMethods
    {
        public static IEnumerable<IVisio.Style> ToEnumerable(this IVisio.Styles styles)
        {
            return CollectionHelpers.ToEnumerable(() => styles.Count, i => styles[i + 1]);
        }

        public static List<IVisio.Style> ToList(this IVisio.Styles styles)
        {
            return CollectionHelpers.ToList(() => styles.Count, i => styles[i + 1]);
        }

        public static string[] GetNamesU(this IVisio.Styles styles)
        {
            System.Array names_sa;
            styles.GetNamesU(out names_sa);
            string[] names = (string[]) names_sa;
            return names;
        }
    }
}