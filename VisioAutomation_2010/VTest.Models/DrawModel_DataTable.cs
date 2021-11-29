using System.Data;
using System.Linq;
using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.Models
{
    public class DrawModel_DataTable : Framework.VTest
    {

        [MUT.TestMethod]
        public void Scripting_Draw_DataTable()
        {
            var pagesize = new VisioAutomation.Core.Size(4, 4);
            var widths = new[] { 2.0, 1.5, 1.0 };
            double default_height = 0.25;
            var cellspacing = new VisioAutomation.Core.Size(0, 0);

            var items = new[]
                {
                    new {Name = "X", Age = 28, Score = 16},
                    new {Name = "Y", Age = 32, Score = 23},
                    new {Name = "Z", Age = 45, Score = 12},
                    new {Name = "U", Age = 48, Score = 10}
                };

            var dt = new DataTable();
            dt.Columns.Add("X", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Score", typeof(int));

            foreach (var item in items)
            {
                dt.Rows.Add(item.Name, item.Age, item.Score);
            }

            // Prepare the Page
            var client = this.GetScriptingClient();
            client.Document.NewDocument();

            var page = client.Page.NewPage(VisioScripting.TargetDocument.Auto, pagesize, false);

            // Draw the table
            var heights = Enumerable.Repeat(default_height, items.Length).ToList();

            var shapes = client.Model.DrawDataTable(VisioScripting.TargetPage.Auto, dt, widths, heights, cellspacing);

            // Verify
            int num_shapes_expected = items.Length*dt.Columns.Count;
            MUT.Assert.AreEqual(num_shapes_expected, shapes.Count);

            // Cleanup
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

    }
}