using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Extensions
{
    public static class MastersMethods
    {
        public static IEnumerable<IVisio.Master> AsEnumerable(this IVisio.Masters masters)
        {
            return masters.Cast<IVisio.Master>();
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