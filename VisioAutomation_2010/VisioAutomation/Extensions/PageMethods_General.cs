using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class PageMethods_General
    {
        public static IEnumerable<IVisio.Page> ToEnumerable(this IVisio.Pages pages)
        {
            return Internal.Extensions.ExtensionHelpers.ToEnumerable(() => pages.Count,
                i => pages[i + 1]);
        }

        public static List<IVisio.Page> ToList(this IVisio.Pages pages)
        {
            return Internal.Extensions.ExtensionHelpers.ToList(() => pages.Count, i => pages[i + 1]);
        }

        public static Core.Rectangle GetBoundingBox(this IVisio.Page page, IVisio.VisBoundingBoxArgs args)
        {
            double bbx0, bby0, bbx1, bby1;
            page.BoundingBox((short) args, out bbx0, out bby0, out bbx1, out bby1);
            var r = new Core.Rectangle(bbx0, bby0, bbx1, bby1);
            return r;
        }

        public static void ResizeToFitContents(this IVisio.Page page, Core.Size padding)
        {
            Pages.PageHelper.ResizeToFitContents(page, padding);
        }


        public static string[] GetNamesU(this IVisio.Pages pages)
        {
            System.Array names_sa;
            pages.GetNamesU(out names_sa);
            string[] names = (string[]) names_sa;
            return names;
        }

    }
}