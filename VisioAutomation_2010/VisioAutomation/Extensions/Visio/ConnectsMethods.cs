using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ConnectsMethods
    {
        public static IEnumerable<IVisio.Connect> AsEnumerable(this IVisio.Connects connects)
        {
            int count = connects.Count;
            for (int i = 0; i < count; i++)
            {
                yield return connects[i+1];
            }
        }
    }
}