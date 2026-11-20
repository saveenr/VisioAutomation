using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page text/tab-stops.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class TextTabStopsSamples
    {
        public static void Sample_02(IVisio.Shape shape)
        {
            // <sample id="text/tab-stops#02" doc="text/tab-stops.md" line="35">
            var stops = new System.Collections.Generic.List<VisioAutomation.Text.TabStop>
            {
                new VisioAutomation.Text.TabStop(1.0, VisioAutomation.Text.TabStopAlignment.Left),
                new VisioAutomation.Text.TabStop(2.5, VisioAutomation.Text.TabStopAlignment.Right),
            };

            VisioAutomation.Text.TextHelper.SetTabStops(shape, stops);
            // </sample>
        }

        public static void Sample_03(IVisio.Shape shape)
        {
            // <sample id="text/tab-stops#03" doc="text/tab-stops.md" line="51">
            var existing = VisioAutomation.Text.TextHelper.GetTabStops(shape);

            foreach (var ts in existing)
            {
                System.Console.WriteLine($"{ts.Position} ({ts.Alignment})");
            }
            // </sample>
        }
    }
}
