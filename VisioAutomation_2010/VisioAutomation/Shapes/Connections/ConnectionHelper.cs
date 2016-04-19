using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Connections
{
    public static class ConnectionHelper
    {
        public static IEnumerable<Connect> ToEnumerable(Microsoft.Office.Interop.Visio.Connects connects)
        {
            int count = connects.Count;
            for (int i = 0; i < count; i++)
            {
                yield return connects[i + 1];
            }
        }
    }
}