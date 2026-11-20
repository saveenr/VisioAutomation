using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/master.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingMasterSamples
    {
        public static void Sample_01(VisioScripting.Client client, int row, double a)
        {
            // <sample id="visio-scripting/master#01" doc="visio-scripting/master.md" line="19">
            // Open a stencil document, then drop one of its masters
            var stencil = client.Document.OpenStencilDocument("basic_u.vss");
            var rectMaster = client.Master.GetMaster(new VisioScripting.TargetDocument(stencil), "Rectangle");

            var s1 = client.Master.DropMaster(VisioScripting.TargetPage.Auto,
                                               rectMaster,
                                               new VisioAutomation.Core.Point(2, 2));

            // Bulk-drop a whole row
            var masters = new[] { rectMaster, rectMaster, rectMaster };
            var points = new[]
            {
                new VisioAutomation.Core.Point(0, 0),
                new VisioAutomation.Core.Point(2, 0),
                new VisioAutomation.Core.Point(4, 0),
            };
            short[] ids = client.Master.DropMasters(VisioScripting.TargetPage.Auto, masters, points);
            // </sample>
        }
    }
}
