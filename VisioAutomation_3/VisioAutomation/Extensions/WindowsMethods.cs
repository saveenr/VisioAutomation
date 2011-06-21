using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Extensions
{
    public static class WindowsMethods
    {
        public static IEnumerable<IVisio.Window> AsEnumerable(this IVisio.Windows windows)
        {
            for (int i = 0; i < windows.Count; i++)
            {
                yield return windows[(short)(i + 1)];
            }
        }
    }
}