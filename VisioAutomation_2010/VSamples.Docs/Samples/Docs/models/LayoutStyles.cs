using IVisio = Microsoft.Office.Interop.Visio;
using VALAY = VisioAutomation.Models.LayoutStyles;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page models/layout-styles.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ModelsLayoutStylesSamples
    {
        public static void Sample_01(IVisio.Page visioPage)
        {
            // <sample id="models/layout-styles#01" doc="models/layout-styles.md" line="40">
            var style = new VALAY.HierarchyLayoutStyle();
            style.LayoutDirection = VALAY.LayoutDirection.TopToBottom;
            style.HorizontalAlignment = VALAY.HorizontalAlignment.Center;
            style.VerticalAlignment = VALAY.VerticalAlignment.Middle;
            style.ConnectorStyle = VALAY.ConnectorStyle.OrganizationChart;
            style.ConnectorAppearance = VALAY.ConnectorAppearance.Default;

            style.Apply(visioPage);
            // </sample>
        }
    }
}
