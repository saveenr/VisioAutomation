using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Extensions
{
    public static class WindowsMethods
    {
        public static IEnumerable<IVisio.Window> AsEnumerable(this IVisio.Windows windows)
        {
            return windows.Cast<IVisio.Window>();
        }
    }
}