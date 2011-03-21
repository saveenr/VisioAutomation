using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomationSamples
{
    public static class PageSamples
    {
        public static void CreateBackgroundPage()
        {
            var bkpage = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            bkpage.Background = 1;
            bkpage.Name = "XBG";
            var s0 = bkpage.DrawRectangle(0, 0, 5, 5);
            s0.Text = "From BK page";

            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            page.Background = 0;
            var s1 = page.DrawRectangle(4, 4, 8, 8);
            s1.Text = "From fg page";
            page.Name = "XFG";

            page.BackPage = bkpage.NameU;
        }
    }
}