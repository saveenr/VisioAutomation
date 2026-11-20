using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page models/introduction.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ModelsIntroductionSamples
    {
        public static void Sample_01(VisioScripting.Client client, double a, System.Data.DataTable dataTable)
        {
            // <sample id="models/introduction#01" doc="models/introduction.md" line="5">
            // 1. Build the model: a table of rows and columns, a tree of people, a graph.
            var model = new VisioAutomation.Models.Data.DataTableModel();
            model.DataTable = dataTable;

            // 2. Render it.
            client.Model.DrawDataTableModel(VisioScripting.TargetPage.Auto, model);
            // </sample>
        }
    }
}
