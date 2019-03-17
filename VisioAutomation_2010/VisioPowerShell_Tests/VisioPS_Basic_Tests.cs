using Microsoft.VisualStudio.TestTools.UnitTesting;
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
        
        private static void VisioPS_Close_Visio_Application()
        {
            bool force = true;
            VisioPS_Basic_Tests.Session.Cmd_Close_VisioApplication(force);
        }

        [TestMethod]
        public void VisioPS_SetShapeCells()
        {
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();
            var stencil_basic = VisioPS_Basic_Tests.Session.Cmd_Open_VisioDocument("basic_u.vss");
            var master_roundrect = VisioPS_Basic_Tests.Session.Cmd_Get_VisioMaster(PsArray.From("Rectangle"), stencil_basic);
            var shapes = VisioPS_Basic_Tests.Session.Cmd_New_VisioShape(PsArray.From(master_roundrect), new [] {2.0, 3.0});

            var shapecells = VisioPS_Basic_Tests.Session.Cmd_New_VisioShapeCells();
            shapecells.XFormPinX= "4 in";
            shapecells.XFormPinY = "6 in";

            VisioPS_Basic_Tests.Session.Cmd_Set_VisioShapeCells(PsArray.From(shapecells), PsArray.From(shapes));

            var dt = VisioPS_Basic_Tests.Session.Cmd_Get_VisioShapeCells(PsArray.From(shapes));

            Assert.IsNotNull(dt);
            Assert.AreEqual("4 in", dt.Rows[0]["XFormPinX"]);
            Assert.AreEqual("6 in", dt.Rows[0]["XFormPinY"]);
            VisioPS_Basic_Tests.Session.Cmd_Close_VisioDocument(PsArray.From(doc), true);
        }

        [TestMethod]
        public void VisioPS_SetPageCells()
        {
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();
            var page = VisioPS_Basic_Tests.Session.Cmd_Get_VisioPage(activepage: true, name: null);

            var pagecells = VisioPS_Basic_Tests.Session.Cmd_New_VisioPageCells();
            pagecells.PageHeight = "4 in";
            pagecells.PageWidth= "3 in";

            VisioPS_Basic_Tests.Session.Cmd_Set_VisioPageCells( PsArray.From(pagecells), PsArray.From(page));
            
            var dt = VisioPS_Basic_Tests.Session.Cmd_Get_VisioPageCells(PsArray.From(page));

            Assert.IsNotNull(dt);
            Assert.AreEqual("3 in", dt.Rows[0]["PageWidth"]);
            Assert.AreEqual("4 in", dt.Rows[0]["PageHeight"]);
            VisioPS_Basic_Tests.Session.Cmd_Close_VisioDocument(PsArray.From(doc), true);
        }

        [TestMethod]
        public void VisioPS_DrawRectangleWithText()
        {
            var d = Session.Cmd_New_VisioDocument();
            var s = Session.Cmd_New_VisioShape(VisioPowerShell.Models.ShapeType.Rectangle, new[] {0.0, 1.0, 2.0, 3.0});
            Session.Cmd_Set_VisioText( PsArray.From("Hello World"), PsArray.From(s));

            var r = Session.Cmd_Get_VisioText();

            Assert.AreEqual(1,r.Length);
            Assert.AreEqual("Hello World", r[0]);
            bool force = true;
            Session.Cmd_Close_VisioDocument( PsArray.From(d), force);
        }

        [TestMethod]
        public void VisioPS_CreateContainer()
        {
            var doc = VisioPS_Basic_Tests.Session.Cmd_New_VisioDocument();
            var app = VisioPS_Basic_Tests.Session.Cmd_Get_VisioApplication();

            var ver = VisioAutomation.Application.ApplicationHelper.GetVersion(app);

            var cont_doc = ver.Major >= 15 ? "SDCONT_U.VSSX" : "SDCONT_U.VSS";
            var cont_master_name = ver.Major >= 15 ? "Plain" : "Container 1";
            var rectangle = "Rectangle";
            var basic_u_vss = "BASIC_U.VSS";

            var master = VisioPS_Basic_Tests.Session.Cmd_Get_VisioMaster(PsArray.From(rectangle), basic_u_vss);


            VisioPS_Basic_Tests.Session.Cmd_New_VisioShape( PsArray.From(master) , new[] { 1.0, 1.0 });

            // Drop a container on the page... the rectangle we created above should be selected by default. 
            // Since it is selected it will be added as a member to the container.

            var container = VisioPS_Basic_Tests.Session.Cmd_New_VisioContainer(cont_master_name, cont_doc);

            Assert.IsNotNull(container);

           

            //VisioPS_Basic_Tests.VisioPS_Close_Visio_Application();
        }
    }
}
