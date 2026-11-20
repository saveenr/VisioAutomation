using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/arrange.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingArrangeSamples
    {
        public static void Sample_01(VisioScripting.Client client)
        {
            // <sample id="visio-scripting/arrange#01" doc="visio-scripting/arrange.md" line="28">
            // Align the current selection to the left edge, then distribute vertically
            client.Arrange.AlignHorizontal(VisioScripting.TargetSelection.Auto,
                                            VisioScripting.Models.AlignmentHorizontal.Left);

            client.Arrange.DistributeOnAxis(VisioScripting.TargetSelection.Auto,
                                             VisioScripting.Models.Axis.YAxis);

            // Nudge the result one quarter inch to the right
            client.Arrange.Nudge(VisioScripting.TargetSelection.Auto, 0.25, 0.0);
            // </sample>
        }
    }
}
