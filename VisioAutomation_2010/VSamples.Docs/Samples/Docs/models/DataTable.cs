using IVisio = Microsoft.Office.Interop.Visio;
using System.Data;
using VAMODELS = VisioAutomation.Models;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page models/data-table.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ModelsDataTableSamples
    {
        public static void Sample_01(VisioScripting.Client client)
        {
            // <sample id="models/data-table#01" doc="models/data-table.md" line="19">
            var dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Role");
            dt.Rows.Add("Alice", "Owner");
            dt.Rows.Add("Bob", "Reviewer");

            var model = new VAMODELS.Data.DataTableModel();
            model.DataTable = dt;
            model.CellWidth = 1.5;
            model.CellHeight = 0.5;
            model.CellSpacing = 0.1;

            client.Model.DrawDataTableModel(VisioScripting.TargetPage.Auto, model);
            // </sample>
        }
    }
}
