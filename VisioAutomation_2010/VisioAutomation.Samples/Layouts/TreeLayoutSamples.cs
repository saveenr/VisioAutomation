using VisioAutomation.Models.Layouts.Tree;
using VADOM = VisioAutomation.Models.Dom;

namespace VisioAutomationSamples
{
    public static class TreeLayoutSamples
    {
        public static void TreeWithTwoPassLayoutAndFormatting()
        {
            var doc = SampleEnvironment.Application.ActiveDocument;
            var page1 = doc.Pages.Add();

            var t = new Drawing();

            t.Root = new Node("Root");

            var na = new Node("A");
            var nb = new Node("B");

            var na1 = new Node("A1");
            var na2 = new Node("A2");

            var nb1 = new Node("B1");
            var nb2 = new Node("B2");

            t.Root.Children.Add(na);
            t.Root.Children.Add(nb);

            na.Children.Add(na1);
            na.Children.Add(na2);

            nb.Children.Add(nb1);
            nb1.Children.Add(nb2);

            var fontname = "Segoe UI";
            var font = doc.Fonts[fontname];

            foreach (var tn in t.Nodes)
            {
                var cells = new VADOM.ShapeCells();
                tn.Cells = cells;

                cells.ParaHorizontalAlign = 0; // align text to left
                cells.TextBlockVerticalAlign = 0; // align text block to top
                cells.CharFont = font.ID;
                cells.CharSize = "10pt";
                cells.FillForeground = "rgb(255,250,200)";
                cells.CharColor = "rgb(255,0,0)";
            }

            t.Render(page1);
        }
    }

}