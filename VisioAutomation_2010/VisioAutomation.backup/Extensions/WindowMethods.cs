﻿using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class WindowMethods
    {
        public static void Select(
            this IVisio.Window window,
            IEnumerable<IVisio.Shape> shapes,
            IVisio.VisSelectArgs selectargs)
        {
            foreach (var shape in shapes)
            {
                window.Select(shape, (short)selectargs);
            }
        }

        public static Geometry.Rectangle GetViewRect(this IVisio.Window window)
        {
            // MSDN: http://msdn.microsoft.com/en-us/library/office/ff765846.aspx
            double left, top, height, width;
            window.GetViewRect(out left, out top, out width, out height);
            double x0 = left;
            double x1 = left + width;
            double y0 = top - height;
            double y1 = y0 + height;

            var r = new Geometry.Rectangle(x0, y0, x1, y1);
            return r;
        }

        public static System.Drawing.Rectangle GetWindowRect(this IVisio.Window window)
        {
            // MSDN: http://msdn.microsoft.com/en-us/library/office/ms367542(v=office.14).aspx
            int left, top, height, width;
            window.GetWindowRect(out left, out top, out width, out height);
            var r = new System.Drawing.Rectangle(left, top, width, height);
            return r;
        }

        public static void SetWindowRect(
            this IVisio.Window window, 
            System.Drawing.Rectangle rect)
        {
            // MSDN: http://msdn.microsoft.com/en-us/library/office/ff769098.aspx
            window.SetWindowRect(rect.Left, rect.Top, rect.Width, rect.Height);
        }

        public static void SetViewRect(
            this IVisio.Window window,
            Geometry.Rectangle rect)
        {
            // MSDN: http://msdn.microsoft.com/en-us/library/office/ms367542(v=office.14).aspx
            window.SetViewRect(rect.Left, rect.Top, rect.Width, rect.Height);
        }

        public static IEnumerable<IVisio.Window> ToEnumerable(this IVisio.Windows windows)
        {
            return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToEnumerable(() => windows.Count, i => windows[(short)(i + 1)]);
        }

        public static List<IVisio.Window> ToList(this IVisio.Windows windows)
        {
            return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToList(() => windows.Count, i => windows[(short)(i + 1)]);
        }
    }
}