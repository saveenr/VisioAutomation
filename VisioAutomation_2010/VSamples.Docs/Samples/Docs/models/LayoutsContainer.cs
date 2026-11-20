using IVisio = Microsoft.Office.Interop.Visio;
using VACONT = VisioAutomation.Models.Layouts.Container;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page models/layouts-container.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ModelsLayoutsContainerSamples
    {
        public static void Sample_01(IVisio.Document visioDoc)
        {
            // <sample id="models/layouts-container#01" doc="models/layouts-container.md" line="17">
            var layout = new VACONT.ContainerLayout();
            var c1 = layout.AddContainer("Fruit");
            c1.Add("Apple");
            c1.Add("Pear");
            var c2 = layout.AddContainer("Vegetables");
            c2.Add("Carrot");

            layout.PerformLayout();
            IVisio.Page page = layout.Render(visioDoc);
            // </sample>
        }
    }
}
