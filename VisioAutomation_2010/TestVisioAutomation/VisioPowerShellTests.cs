using Microsoft.VisualStudio.TestTools.UnitTesting;
using SMA = System.Management.Automation;

namespace TestVisioAutomation
{
    [TestClass]
    public class VisioPowerShellTests
    {
        private static readonly VisioPSContext visiops_session = new VisioPSContext();

        [ClassInitialize]
        public static void PSTestFixtureSetup(TestContext context)
        {
            var new_visio_application = new VisioPowerShell.Commands.New.New_VisioApplication();
            var visio_app = new_visio_application.Invoke();
        }

        [TestCleanup]
        public void PSTestFixtureTeardown()
        {
           
        }
 
        [ClassCleanup]
        public static void CleanUp()
        {
            VisioPowerShellTests.visiops_session.CleanUp();
        }
 
        [TestMethod]
        public void VisioPS_New_Visio_Document()
        {
            var doc = VisioPowerShellTests.visiops_session.New_Visio_Document();
            Assert.IsNotNull(doc);
            VisioPowerShellTests.Close_Visio_Application();
        }

        private static void Close_Visio_Application()
        {
            VisioPowerShellTests.visiops_session.Close_Visio_Application();
        }

        [TestMethod]
        public void VisioPS_Set_Visio_Page_Cell()
        {

            // Handle the page that gets created when a document is created

            var doc = VisioPowerShellTests.visiops_session.New_Visio_Document();
            var dic = new System.Collections.Generic.Dictionary<string, object>
            {
                {"PageWidth", 3},
                {"PageHeight", 5}
            };

            VisioPowerShellTests.visiops_session.Set_Visio_PageCells(dic);

            //VisioPowerShellTests.Close_Visio_Application();
        }

        [TestMethod]
        public void VisioPS_Get_Visio_Page_Cell()
        {
            var cells = new[] { "PageWidth", "PageHeight" };
            var result_type = "Double";
            var get_results = true;

            // Handle the page that gets created when a document is created

            var doc = VisioPowerShellTests.visiops_session.New_Visio_Document();
            var datatable1 = VisioPowerShellTests.visiops_session.Get_Visio_Page_Cell(cells, get_results, result_type);

            Assert.IsNotNull(datatable1);
            Assert.AreEqual(8.5, datatable1.Rows[0]["PageWidth"]);
            Assert.AreEqual(11.0, datatable1.Rows[0]["PageHeight"]);
            
            //Now lets add another page and get it's width and height
            var page2 = VisioPowerShellTests.visiops_session.New_Visio_Page();
            var datatable2 = VisioPowerShellTests.visiops_session.Get_Visio_Page_Cell(cells, get_results, result_type);
 
            Assert.IsNotNull(datatable2);
            Assert.AreEqual(8.5, datatable2.Rows[0]["PageWidth"]);
            Assert.AreEqual(11.0, datatable2.Rows[0]["PageHeight"]);

            VisioPowerShellTests.Close_Visio_Application();
        }


      [TestMethod]
      public void VisioPS_NewVisioContainer()
      {
          var doc = VisioPowerShellTests.visiops_session.New_Visio_Document();
          var app = VisioPowerShellTests.visiops_session.Get_Visio_Application();

          var ver = VisioAutomation.Application.ApplicationHelper.GetVersion(app);

          var cont_doc = ver.Major >= 15 ? "SDCONT_U.VSSX" : "SDCONT_U.VSS";
          var cont_master_name = ver.Major >= 15 ? "Plain" : "Container 1";
          var rectangle = "Rectangle";
          var basic_u_vss = "BASIC_U.VSS";

          var rect = VisioPowerShellTests.visiops_session.Get_Visio_Master(rectangle, basic_u_vss);

          VisioPowerShellTests.visiops_session.New_VisioShape(rect, new[] { 1.0, 1.0 });

          // Drop a container on the page... the rectangle we created above should be selected by default. 
          // Since it is selected it will be added as a member to the container.

          var container = VisioPowerShellTests.visiops_session.New_Visio_Container(cont_master_name, cont_doc);

          Assert.IsNotNull(container);

          VisioPowerShellTests.Close_Visio_Application();
      }

    }

    public static class SMA_Extensions
    {
        public static void AddParameter(this SMA.Runspaces.Command cmd, string name, object value)
        {
            var parameter= new SMA.Runspaces.CommandParameter(name, value);
            cmd.Parameters.Add(parameter);            
        }    
    }
}
	 