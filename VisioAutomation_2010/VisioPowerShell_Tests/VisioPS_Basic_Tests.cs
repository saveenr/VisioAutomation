using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioPowerShell.Models;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioPowerShell_Tests.Framework;

namespace VisioPowerShell_Tests
{
    [TestClass]
    public class VisioPS_Basic_Tests
    {
        private static readonly VisioPS_Session Session = new VisioPS_Session();

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            var new_visio_application = new VisioPowerShell.Commands.NewVisioApplication();
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            VisioPS_Basic_Tests.Session.CleanUp();
        }
        
        [TestMethod]
        public void VisioPS_New_Visio_Document()
        {
            var doc = VisioPS_Basic_Tests.Session.New_VisioDocument();
            Assert.IsNotNull(doc);
            VisioPS_Basic_Tests.VisioPS_Close_Visio_Application();
        }

        private static void VisioPS_Close_Visio_Application()
        {
            VisioPS_Basic_Tests.Session.Close_VisioApplication();
        }

        [TestMethod]
        public void VisioPS_Set_Visio_Shape_Cell()
        {
            var doc = VisioPS_Basic_Tests.Session.New_VisioDocument();
            var basic_stencil = VisioPS_Basic_Tests.Session.Open_VisioDocument("basic_u.vss");
            var rrect_master = VisioPS_Basic_Tests.Session.Get_VisioMaster("Rectangle", basic_stencil);
            var shapes = VisioPS_Basic_Tests.Session.New_VisioShape(PsArray.From(rrect_master), new [] {2.0, 3.0});

            var cells = VisioPS_Basic_Tests.Session.New_VisioShapeCells();
            cells.XFormPinX= "4 in";
            cells.XFormPinY = "6 in";

            VisioPS_Basic_Tests.Session.Set_VisioShapeCells(PsArray.From(cells), PsArray.From(shapes));

            var dt = VisioPS_Basic_Tests.Session.Get_VisioShapeCells(PsArray.From(shapes));

            Assert.IsNotNull(dt);
            Assert.AreEqual("4 in", dt.Rows[0]["XFormPinX"]);
            Assert.AreEqual("6 in", dt.Rows[0]["XFormPinY"]);
            VisioPS_Basic_Tests.Session.Close_VisioDocument(PsArray.From(doc), true);
        }

        [TestMethod]
        public void VisioPS_Set_Visio_Page_Cell()
        {
            var doc = VisioPS_Basic_Tests.Session.New_VisioDocument();
            var page = VisioPS_Basic_Tests.Session.Get_VisioPage(activepage: true, name: null);

            var cells = VisioPS_Basic_Tests.Session.New_VisioPageCells();
            var pagecells = cells;
            pagecells.PageHeight = "4 in";
            pagecells.PageWidth= "3 in";

            VisioPS_Basic_Tests.Session.Set_VisioPageCells( PsArray.From(cells), PsArray.From(page));
            
            var dt = VisioPS_Basic_Tests.Session.Get_VisioPageCells(PsArray.From(page));

            Assert.IsNotNull(dt);
            Assert.AreEqual("3 in", dt.Rows[0]["PageWidth"]);
            Assert.AreEqual("4 in", dt.Rows[0]["PageHeight"]);
            VisioPS_Basic_Tests.Session.Close_VisioDocument(PsArray.From(doc), true);
        }

        [TestMethod]
        public void VisioPS_DrawRect()
        {
            var d = Session.New_VisioDocument();
            var s = Session.New_VisioShape(VisioPowerShell.Commands.ShapeType.Rectangle, new[] {0.0, 1.0, 2.0, 3.0});
            Session.Set_VisioShapeText( PsArray.From("Hello World"), PsArray.From(s));

            var r = Session.Get_VisioShapeText();

            Assert.AreEqual(1,r.Length);
            Assert.AreEqual("Hello World", r[0]);
            bool force = true;
            Session.Close_VisioDocument( PsArray.From(d), force);
        }

        [TestMethod]
        public void VisioPS_NewVisioContainer()
        {
            var doc = VisioPS_Basic_Tests.Session.New_VisioDocument();
            var app = VisioPS_Basic_Tests.Session.Get_VisioApplication();

            var ver = VisioAutomation.Application.ApplicationHelper.GetVersion(app);

            var cont_doc = ver.Major >= 15 ? "SDCONT_U.VSSX" : "SDCONT_U.VSS";
            var cont_master_name = ver.Major >= 15 ? "Plain" : "Container 1";
            var rectangle = "Rectangle";
            var basic_u_vss = "BASIC_U.VSS";

            var masters = VisioPS_Basic_Tests.Session.Get_VisioMaster(rectangle, basic_u_vss);

            VisioPS_Basic_Tests.Session.New_VisioShape( PsArray.From(masters) , new[] { 1.0, 1.0 });

            // Drop a container on the page... the rectangle we created above should be selected by default. 
            // Since it is selected it will be added as a member to the container.

            var container = VisioPS_Basic_Tests.Session.New_VisioContainer(cont_master_name, cont_doc);

            Assert.IsNotNull(container);

            VisioPS_Basic_Tests.VisioPS_Close_Visio_Application();
        }
    }
}
