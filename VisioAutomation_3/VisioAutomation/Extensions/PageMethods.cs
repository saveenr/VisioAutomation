using System;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Extensions
{
    public static partial class PageMethods
    {
        public static void Activate(this IVisio.Page page)
        {
            VA.Pages.PageHelper.Activate(page);
        }

        public static void ResizeToFitContents(this IVisio.Page page, double borderwidth, double borderheight)
        {
            var bordersize = new VA.Drawing.Size(borderwidth, borderheight);
            VA.Pages.PageHelper.ResizeToFitContents(page, bordersize);
        }

        public static void ResizeToFitContents(this IVisio.Page page, VA.Drawing.Size bordersize)
        {
            VA.Pages.PageHelper.ResizeToFitContents(page, bordersize);
        }

        public static VA.Drawing.Size GetSize(this IVisio.Page page)
        {
            return VA.Pages.PageHelper.GetSize(page);
        }

        public static void SetSize(this IVisio.Page page, VA.Drawing.Size size)
        {
            VA.Pages.PageHelper.SetSize(page, size);
        }

        public static void SetSize(this IVisio.Page page, double x, double y)
        {
            VA.Pages.PageHelper.SetSize(page, new VA.Drawing.Size(x, y));
        }
    }
}