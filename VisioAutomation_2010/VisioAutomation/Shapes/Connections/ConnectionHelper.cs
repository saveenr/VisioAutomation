using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Connections
{
    public static class ConnectionHelper
    {
        public static IEnumerable<IVisio.Connect> ToEnumerable(IVisio.Connects connects)
        {
            int count = connects.Count;
            for (int i = 0; i < count; i++)
            {
                yield return connects[i + 1];
            }
        }
    }
}