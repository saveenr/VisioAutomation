using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioInterop
{
    public static class Util
    {
        public static IVisio.Shape CreateStandardShape(IVisio.Page page)
        {
            var shape = page.DrawRectangle(1, 1, 4, 3);
            var cell_width = shape.CellsU["Width"];
            var cell_height = shape.CellsU["Height"];
            cell_width.Formula = "=(1.0+2.5)";
            cell_height.Formula = "=(0.0+1.5)";
            return shape;
        }

        public static IVisio.Page CreateStandardPage(IVisio.Document doc, string pagename)
        {
            var pages = doc.Pages;
            var page = pages.Add();
            page.NameU = pagename;
            return page;
        }
    }
}