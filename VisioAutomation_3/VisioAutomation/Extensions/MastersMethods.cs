using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Extensions
{
    public static class MastersMethods
    {
        public static IEnumerable<IVisio.Master> AsEnumerable(this IVisio.Masters masters)
        {
            for (int i = 0; i < masters.Count; i++)
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