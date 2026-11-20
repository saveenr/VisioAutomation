using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/model.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingModelSamples
    {
        public static void Sample_01(VisioScripting.Client client)
        {
            // <sample id="visio-scripting/model#01" doc="visio-scripting/model.md" line="34">
            // Render a small table on the active page
            var dt = new System.Data.DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Role");
            dt.Rows.Add("Alice", "Owner");
            dt.Rows.Add("Bob", "Reviewer");

            var model = new VisioAutomation.Models.Data.DataTableModel();
            model.DataTable = dt;
            model.CellWidth = 1.5;     // inches
            model.CellHeight = 0.5;
            client.Model.DrawDataTableModelIntoPage(VisioScripting.TargetPage.Auto, model);

            // Load a directed-graph XML and render it onto fresh pages
            var dgxml = System.Xml.Linq.XDocument.Load(@"C:\models\graph.xml");
            var dgdoc = client.Model.LoadDirectedGraphFromXml(dgxml);
            client.Model.DrawDirectedGraphDocument(dgdoc, new VisioAutomation.Models.Layouts.DirectedGraph.DirectedGraphStyling());
            // </sample>
        }
    }
}
