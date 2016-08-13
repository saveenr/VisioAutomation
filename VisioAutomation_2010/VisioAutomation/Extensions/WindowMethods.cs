using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class WindowMethods
    {
        public static void Select(
            this Microsoft.Office.Interop.Visio.Window window,
            IEnumerable<Shape> shapes,
            Microsoft.Office.Interop.Visio.VisSelectArgs selectargs)
        {
            VisioAutomation.Windows.WindowHelper.Select(window, shapes, selectargs);
        }

        public static Drawing.Rectangle GetViewRect(this Microsoft.Office.Interop.Visio.Window window)
        {
            return VisioAutomation.Windows.WindowHelper.GetViewRect(window);
        }

        public static System.Drawing.Rectangle GetWindowRect(this Microsoft.Office.Interop.Visio.Window window)
        {
            return VisioAutomation.Windows.WindowHelper.GetWindowRect(window);
        }

        public static void SetWindowRect(
            this Microsoft.Office.Interop.Visio.Window window, 
            System.Drawing.Rectangle rect)
        {
            VisioAutomation.Windows.WindowHelper.SetWindowRect(window, rect);
        }

        public static void SetViewRect(
            this Microsoft.Office.Interop.Visio.Window window, 
            Drawing.Rectangle rect)
        {
            VisioAutomation.Windows.WindowHelper.SetViewRect(window,rect);
        }

        public static IEnumerable<Window> ToEnumerable(this Microsoft.Office.Interop.Visio.Windows windows)
        {
            return VisioAutomation.Windows.WindowHelper.ToEnumerable(windows);
        }
    }
}