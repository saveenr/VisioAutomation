using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class PagesMethods
    {
        public static IEnumerable<IVisio.Page> AsEnumerable(this IVisio.Pages pages)
        {
            short count = pages.Count;
            for (int i = 0; i < count; i++)
            {
                yield return pages[i + 1];
            }
        }

        public static string[] GetNamesU(this IVisio.Pages pages)
        {
            System.Array names_sa;
            pages.GetNamesU(out names_sa);
            string[] names = (string[]) names_sa;
            return names;
        }
    }
}