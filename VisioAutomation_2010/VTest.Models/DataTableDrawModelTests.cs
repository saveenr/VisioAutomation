using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.Models
{
    [MUT.TestClass]
    public class DataTableDrawModelTests : Framework.VTest
    {

        [MUT.TestMethod]
        public void RenderDataTable_FromSampleData_DoesNotThrow()
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

            var dt = new System.Data.DataTable();
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

        [MUT.TestMethod]
        public void DrawDataTable_WidthsAndHeights_AreHonored()
        {
            var dt = make_table(2, 2);
            var client = this.GetScriptingClient();
            client.Document.NewDocument();

            var shapes = client.Model.DrawDataTable(
                VisioScripting.TargetPage.Auto,
                dt,
                new[] { 2.0, 1.5 },
                new[] { 0.5, 0.25 },
                new VisioAutomation.Core.Size(0, 0));

            assert_size(shapes, "r0c0", 2.0, 0.5);
            assert_size(shapes, "r0c1", 1.5, 0.5);
            assert_size(shapes, "r1c0", 2.0, 0.25);
            assert_size(shapes, "r1c1", 1.5, 0.25);

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void DrawDataTable_ShorterWidthsList_LeavesRemainingColumnsAtDefaultSize()
        {
            var dt = make_table(3, 1);
            var client = this.GetScriptingClient();
            client.Document.NewDocument();

            var shapes = client.Model.DrawDataTable(
                VisioScripting.TargetPage.Auto,
                dt,
                new[] { 2.0 },
                new[] { 0.5 },
                new VisioAutomation.Core.Size(0, 0));

            assert_size(shapes, "r0c0", 2.0, 0.5);
            assert_size(shapes, "r0c1", 1.0, 0.5);
            assert_size(shapes, "r0c2", 1.0, 0.5);

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void DrawDataTable_NonPositiveWidth_ThrowsArgumentOutOfRangeException()
        {
            var dt = make_table(2, 1);
            var client = this.GetScriptingClient();
            client.Document.NewDocument();

            MUT.Assert.ThrowsExactly<System.ArgumentOutOfRangeException>(
                () => client.Model.DrawDataTable(
                    VisioScripting.TargetPage.Auto,
                    dt,
                    new[] { 0.0, 1.0 },
                    new[] { 0.5 },
                    new VisioAutomation.Core.Size(0, 0)));

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void DrawDataTableModel_CellWidthAndHeight_AreHonored()
        {
            var model = new VisioAutomation.Models.Data.DataTableModel();
            model.DataTable = make_table(2, 2);
            model.CellWidth = 3.0;
            model.CellHeight = 0.5;
            model.CellSpacing = 0.1;

            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            client.Model.DrawDataTableModel(VisioScripting.TargetPage.Auto, model);

            var shapes = this.GetVisioApplication().ActivePage.Shapes.Cast<IVisio.Shape>().ToList();
            MUT.Assert.AreEqual(4, shapes.Count);
            foreach (var shape in shapes)
            {
                MUT.Assert.AreEqual(3.0, shape.Cells["Width"].ResultIU, 1e-6);
                MUT.Assert.AreEqual(0.5, shape.Cells["Height"].ResultIU, 1e-6);
            }

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void DrawDataTableModel_CellSizeNotSet_DrawsOneByOneInchCells()
        {
            var model = new VisioAutomation.Models.Data.DataTableModel();
            model.DataTable = make_table(2, 1);

            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            client.Model.DrawDataTableModel(VisioScripting.TargetPage.Auto, model);

            var shapes = this.GetVisioApplication().ActivePage.Shapes.Cast<IVisio.Shape>().ToList();
            MUT.Assert.AreEqual(2, shapes.Count);
            foreach (var shape in shapes)
            {
                MUT.Assert.AreEqual(1.0, shape.Cells["Width"].ResultIU, 1e-6);
                MUT.Assert.AreEqual(1.0, shape.Cells["Height"].ResultIU, 1e-6);
            }

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void DrawDataTableModel_DrawsOnTheTargetPage_NotTheActivePage()
        {
            var model = new VisioAutomation.Models.Data.DataTableModel();
            model.DataTable = make_table(2, 2);

            var app = this.GetVisioApplication();
            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            var page_a = app.ActivePage;
            var page_b = client.Page.NewPage(VisioScripting.TargetDocument.Auto, new VisioAutomation.Core.Size(8.5, 11), false);
            app.ActiveWindow.Page = page_a;

            client.Model.DrawDataTableModel(new VisioScripting.TargetPage(page_b), model);

            MUT.Assert.AreEqual(4, page_b.Shapes.Count, "the table should be on the target page");
            MUT.Assert.AreEqual(0, page_a.Shapes.Count, "the active page should be untouched");

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        private static System.Data.DataTable make_table(int columns, int rows)
        {
            var dt = new System.Data.DataTable();
            foreach (int c in Enumerable.Range(0, columns))
            {
                dt.Columns.Add("c" + c, typeof(string));
            }

            foreach (int r in Enumerable.Range(0, rows))
            {
                var values = Enumerable.Range(0, columns).Select(c => (object)("r" + r + "c" + c)).ToArray();
                dt.Rows.Add(values);
            }

            return dt;
        }

        private static void assert_size(System.Collections.Generic.List<IVisio.Shape> shapes, string text, double width, double height)
        {
            var shape = shapes.Single(s => s.Text == text);
            MUT.Assert.AreEqual(width, shape.Cells["Width"].ResultIU, 1e-6, text + " width");
            MUT.Assert.AreEqual(height, shape.Cells["Height"].ResultIU, 1e-6, text + " height");
        }

    }
}