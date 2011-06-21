using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Extensions
{
    public static class ConnectsMethods
    {
        public static IEnumerable<IVisio.Connect> AsEnumerable(this IVisio.Connects connects)
        {
            for (int i = 0; i < connects.Count; i++)
            {
                yield return connects[i+1];
            }
        }
    }
}