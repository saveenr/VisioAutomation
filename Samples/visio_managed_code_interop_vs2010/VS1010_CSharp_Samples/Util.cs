using Microsoft.Office.Interop.Visio;
using IVisio = Microsoft.Office.Interop.Visio;

public static class Util
{
    public static Shape CreateStandardShape(Page page)
    {
        var shape = page.DrawRectangle(1, 1, 4, 3);
        var cell_width = shape.CellsU["Width"];
        var cell_height = shape.CellsU["Height"];
        cell_width.Formula = "=(1.0+2.5)";
        cell_height.Formula = "=(0.0+1.5)";
        return shape;
    }

    public static Page CreateStandardPage(Document doc, string pagename)
    {
        var pages = doc.Pages;
        var page = pages.Add();
        page.NameU = pagename;
        return page;
    }
}
