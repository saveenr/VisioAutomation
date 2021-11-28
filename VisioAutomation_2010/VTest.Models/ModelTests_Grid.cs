using VisioAutomation.Extensions;
using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;
using GRID = VisioAutomation.Models.Layouts.Grid;

namespace VTest.Models
{
    public class ModelTests_Grid : Framework.VTest
    {
        [MUT.TestMethod]
        public void Scripting_Draw_Grid()
        {
            var origin = new VisioAutomation.Core.Point(0, 4);
            var pagesize = new VisioAutomation.Core.Size(4, 4);
            var cellsize = new VisioAutomation.Core.Size(0.5, 0.25);
            int cols = 3;
            int rows = 6;

            // Create the Page
            var client = this.GetScriptingClient();
            client.Document.NewDocument();

            client.Page.NewPage(VisioScripting.TargetDocument.Auto, pagesize, false);

            // Find the stencil and master
            var stencildoc = client.Document.OpenStencilDocument("basic_u.vss");
            var stencil_targetdoc = new VisioScripting.TargetDocument(stencildoc);
            var master = client.Master.GetMaster(stencil_targetdoc, "Rectangle");

            // Draw the grid
            var page = client.Page.GetActivePage();
            var grid = new GRID.GridLayout(cols, rows, cellsize, master);
            grid.Origin = origin;
            grid.Render(page);

            // Verify
            int total_shapes_expected = cols*rows;
            var shapes = page.Shapes.ToList();
            int total_shapes_actual = shapes.Count;
            MUT.Assert.AreEqual(total_shapes_expected,total_shapes_actual);

            // Cleanup
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

    }
}