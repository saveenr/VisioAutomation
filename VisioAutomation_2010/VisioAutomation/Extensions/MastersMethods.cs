using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class MastersMethods
    {
        public static IEnumerable<IVisio.Master> AsEnumerable(this IVisio.Masters masters)
        {
            short count = masters.Count;
            for (int i = 0; i < count; i++)
            {
                yield return masters[i + 1];
            }
        }

        public static string[] GetNamesU(this IVisio.Masters masters)
        {
            System.Array names_sa;
            masters.GetNamesU(out names_sa);
            string[] names = (string[])names_sa;
            return names;
        }
    }
}