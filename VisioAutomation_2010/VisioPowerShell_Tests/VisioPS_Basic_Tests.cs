using Microsoft.VisualStudio.TestTools.UnitTesting;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioPowerShell_Tests.Framework;

namespace VisioPowerShell_Tests
{
    [TestClass]
    public class VisioPS_Basic_Tests
    {
        private static readonly VisioPS_Session session = new VisioPS_Session();

        [ClassInitialize]
        public static void PSTestFixtureSetup(TestContext context)
        {
            var new_visio_application = new VisioPowerShell.Commands.NewVisioApplication();
        }

        [TestCleanup]
        public void PSTestFixtureTeardown()
        {

        }

        [ClassCleanup]
        public static void CleanUp()
        {
            VisioPS_Basic_Tests.session.CleanUp();
        }

        [TestMethod]
        public void VisioPS_New_Visio_Document()
        {
            var doc = VisioPS_Basic_Tests.session.New_VisioDocument();
            Assert.IsNotNull(doc);
            VisioPS_Basic_Tests.Close_Visio_Application();
        }

        private static void Close_Visio_Application()
        {
            VisioPS_Basic_Tests.session.Close_VisioApplication();
        }

        [TestMethod]
        public void VisioPS_Set_Visio_Page_Cell()
        {
            /*
            // Handle the page that gets created when a document is created

            var doc = VisioPowerShellTests.visiops_session.New_Visio_Document();
            var dic = new System.Collections.Generic.Dictionary<string, object>
            {
                {"PageWidth", 3},
                {"PageHeight", 5}
            };

            VisioPowerShellTests.visiops_session.Set_Visio_PageCells(dic);

            //VisioPowerShellTests.Close_Visio_Application();
            */
        }

        [TestMethod]
        public void VisioPS_DrawRect()
        {
            var d = session.New_VisioDocument();
            var s = session.New_VisioShape(VisioPowerShell.Commands.ShapeType.Rectangle, new[] {0.0, 1.0, 2.0, 3.0});
            session.Set_VisioShapeText( PsArray.From("Hello World"), PsArray.From(s));

            var r = session.Get_VisioShapeText();

            Assert.AreEqual(1,r.Length);
            Assert.AreEqual("Hello World", r[0]);
            bool force = true;
            session.Close_VisioDocument( PsArray.From(d), force);
        }

        [TestMethod]
        public void VisioPS_Get_Visio_Page_Cell()
        {
            var doc = VisioPS_Basic_Tests.session.New_VisioDocument();
            var page = VisioPS_Basic_Tests.session.Get_VisioPage(active:true,name:null);

            var datatable1 = VisioPS_Basic_Tests.session.Get_VisioShapeCells( PsArray.From(page.PageSheet));

            Assert.IsNotNull(datatable1);
            Assert.AreEqual("8.5 in", datatable1.Rows[0]["PageWidth"]);
            Assert.AreEqual("11 in", datatable1.Rows[0]["PageHeight"]);

            /*
                
                //Now lets add another page and get it's width and height
                var page2 = VisioPowerShellTests.visiops_session.New_Visio_Page();
                var datatable2 = VisioPowerShellTests.visiops_session.Get_Visio_Page_Cell(cells, get_results, result_type);
     
                Assert.IsNotNull(datatable2);
                Assert.AreEqual(8.5, datatable2.Rows[0]["PageWidth"]);
                Assert.AreEqual(11.0, datatable2.Rows[0]["PageHeight"]);
    
                VisioPowerShellTests.Close_Visio_Application();
            */
        }

        [TestMethod]
        public void VisioPS_NewVisioContainer()
        {
            var doc = VisioPS_Basic_Tests.session.New_VisioDocument();
            var app = VisioPS_Basic_Tests.session.Get_VisioApplication();

            var ver = VisioAutomation.Application.ApplicationHelper.GetVersion(app);

            var cont_doc = ver.Major >= 15 ? "SDCONT_U.VSSX" : "SDCONT_U.VSS";
            var cont_master_name = ver.Major >= 15 ? "Plain" : "Container 1";
            var rectangle = "Rectangle";
            var basic_u_vss = "BASIC_U.VSS";

            var rect = VisioPS_Basic_Tests.session.Get_VisioMaster(rectangle, basic_u_vss);

            VisioPS_Basic_Tests.session.New_VisioShape( PsArray.From(rect) , new[] { 1.0, 1.0 });

            // Drop a container on the page... the rectangle we created above should be selected by default. 
            // Since it is selected it will be added as a member to the container.

            var container = VisioPS_Basic_Tests.session.New_VisioContainer(cont_master_name, cont_doc);

            Assert.IsNotNull(container);

            VisioPS_Basic_Tests.Close_Visio_Application();
        }
    }
}
