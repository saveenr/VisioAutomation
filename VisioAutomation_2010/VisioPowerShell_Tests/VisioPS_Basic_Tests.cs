using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioPowerShell.Models;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioPowerShell_Tests.Framework;

namespace VisioPowerShell_Tests
{
    [TestClass]
    public class VisioPS_Basic_Tests
    {
        private static readonly VisioPS_Session session = new VisioPS_Session();

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            var new_visio_application = new VisioPowerShell.Commands.NewVisioApplication();
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            VisioPS_Basic_Tests.session.CleanUp();
        }
        
        [TestMethod]
        public void VisioPS_New_Visio_Document()
        {
            var doc = VisioPS_Basic_Tests.session.New_VisioDocument();
            Assert.IsNotNull(doc);
            VisioPS_Basic_Tests.VisioPS_Close_Visio_Application();
        }

        private static void VisioPS_Close_Visio_Application()
        {
            VisioPS_Basic_Tests.session.Close_VisioApplication();
        }

        [TestMethod]
        public void VisioPS_Set_Visio_Page_Cell()
        {
            var doc = VisioPS_Basic_Tests.session.New_VisioDocument();
            var page = VisioPS_Basic_Tests.session.Get_VisioPage(activepage: true, name: null);

            var cells = VisioPS_Basic_Tests.session.New_VisioPageCells();
            var pagecells = cells;
            pagecells.PageHeight = "4 in";
            pagecells.PageWidth= "3 in";

            VisioPS_Basic_Tests.session.Set_VisioPageCells(cells, page);
            
            var datatable1 = VisioPS_Basic_Tests.session.Get_VisioPageCells(page);

            Assert.IsNotNull(datatable1);
            Assert.AreEqual("3 in", datatable1.Rows[0]["PageWidth"]);
            Assert.AreEqual("4 in", datatable1.Rows[0]["PageHeight"]);
            VisioPS_Basic_Tests.session.Close_VisioDocument(PsArray.From(doc), true);
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
            var page = VisioPS_Basic_Tests.session.Get_VisioPage(activepage:true,name:null);

            var datatable1 = VisioPS_Basic_Tests.session.Get_VisioPageCells( page );

            Assert.IsNotNull(datatable1);
            Assert.AreEqual("8.5 in", datatable1.Rows[0]["PageWidth"]);
            Assert.AreEqual("11 in", datatable1.Rows[0]["PageHeight"]);
            VisioPS_Basic_Tests.session.Close_VisioDocument(PsArray.From(doc),true);
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

            VisioPS_Basic_Tests.session.New_VisioShape( rect.ToArray() , new[] { 1.0, 1.0 });

            // Drop a container on the page... the rectangle we created above should be selected by default. 
            // Since it is selected it will be added as a member to the container.

            var container = VisioPS_Basic_Tests.session.New_VisioContainer(cont_master_name, cont_doc);

            Assert.IsNotNull(container);

            VisioPS_Basic_Tests.VisioPS_Close_Visio_Application();
        }
    }
}
