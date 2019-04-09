using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingModelTests : VisioAutomationTest
    {
        [TestMethod]
        public void DrawGrid()
        {
            var client = this.GetScriptingClient();
            var page_size = new VisioAutomation.Geometry.Size(10,5);

            var doc = client.Document.NewDocument(page_size);

            var stencildoc = client.Document.OpenStencilDocument("basic_u.vss");


            var rectangle_master = client.Master.GetMaster(new VisioScripting.TargetDocument(stencildoc), "Rectangle");


            var grid = new VisioAutomation.Models.Layouts.Grid.GridLayout(3,2, new VisioAutomation.Geometry.Size(1,2), rectangle_master);

            var page1 = client.Page.NewPage(null, false);
            var page2 = client.Page.NewPage(null, false);
            var page3 = client.Page.NewPage(null, false);

            var target_page = new VisioScripting.TargetPage(page2);
            client.Model.DrawGrid(target_page,grid);
            client.Page.ResizePageToFitContents(target_page, new VisioAutomation.Geometry.Size(0,0));


            client.Document.CloseDocument(new VisioScripting.TargetDocument(), true);
        }
    }
}