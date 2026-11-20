using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VATREE = VisioAutomation.Models.Layouts.Tree;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page models/layouts-tree.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ModelsLayoutsTreeSamples
    {
        public static void Sample_01(IVisio.Page visioPage)
        {
            // <sample id="models/layouts-tree#01" doc="models/layouts-tree.md" line="17">
            var t = new VATREE.Drawing();
            t.Root = new VATREE.Node("Root");

            var na = new VATREE.Node("A");
            var nb = new VATREE.Node("B");
            var na1 = new VATREE.Node("A1");
            var na2 = new VATREE.Node("A2");

            t.Root.Children.Add(na);
            t.Root.Children.Add(nb);
            na.Children.Add(na1);
            na.Children.Add(na2);

            t.LayoutOptions.DefaultNodeSize = new VA.Core.Size(1, 1);
            t.Render(visioPage);
            // </sample>
        }
    }
}
