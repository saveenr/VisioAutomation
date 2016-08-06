using VA = VisioAutomation;
using TREEMODEL = VisioAutomation.Models.Tree;

namespace VisioAutomationSamples
{
    public static class TreeLayoutSamples
    {
        public static void TreeWithTwoPassLayoutAndFormatting()
        {
            var doc = SampleEnvironment.Application.ActiveDocument;
            var page1 = doc.Pages.Add();

            var t = new TREEMODEL.Drawing();

            t.Root = new TREEMODEL.Node("Root");

            var na = new TREEMODEL.Node("A");
            var nb = new TREEMODEL.Node("B");

            var na1 = new TREEMODEL.Node("A1");
            var na2 = new TREEMODEL.Node("A2");

            var nb1 = new TREEMODEL.Node("B1");
            var nb2 = new TREEMODEL.Node("B2");

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
                var cells = new VA.DOM.ShapeCells();
                tn.Cells = cells;

                cells.ParaHorizontalAlign = 0; // align text to left
                cells.VerticalAlign = 0; // align text block to top
                cells.CharFont = font.ID;
                cells.CharSize = "10pt";
                cells.FillForegnd = "rgb(255,250,200)";
                cells.CharColor = "rgb(255,0,0)";
            }

            // TODO: Complete this sample
        }
    }

}