using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class WindowsMethods
    {
        public static IEnumerable<IVisio.Window> AsEnumerable(this IVisio.Windows windows)
        {
            short count = windows.Count;
            for (int i = 0; i < count; i++)
            {
                yield return windows[(short)(i + 1)];
            }
        }
    }
}